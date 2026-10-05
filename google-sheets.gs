/**
 * 拍照盤點 → Google 試算表
 * 用法（只要做一次）：
 * 1. 開一個新的 Google 試算表 → 上方「擴充功能」→「Apps Script」
 * 2. 把這整段貼上（取代原本的內容）→ 存檔
 * 3. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    執行身分：我　／　誰可以存取：所有人 → 部署 → 第一次會要你授權（選自己的帳號 → 進階 → 前往）
 * 4. 複製「網頁應用程式網址」（https://script.google.com/macros/s/…/exec），貼到拍照盤點 App 的「設定 → Google 試算表」
 * 之後每次在 App 按「完成」，這次的盤點就會寫進來；同一次盤點重複送，會先刪掉舊的再寫，不會重複。
 *
 * 會自動建立四個工作表：
 * - 盤點紀錄：每一次、每一種商品一列（原始資料，不要手動改欄位順序）
 * - 總表：每個品項在每個盤點日期各幾件（自動計算）
 * - 最新一次：最近一次盤點的品項與數量，數量多的排前面（自動計算）
 * - 品項庫：料號、實盤、帳面、差異、剩幾個要叫貨、該叫貨（每次同步整張更新，不要手動改）
 *
 * 多人同步（手機、電腦、同事看到同一份資料）：
 * - 盤點紀錄（含照片）、品項庫、儲位、樣品照存在你自己的 Google 雲端硬碟「拍照盤點同步資料（不要刪）」資料夾。
 * - 權限像 Google 雲端硬碟的「共用」：擁有者（第一台）、管理員（可以邀請／移除人）、編輯者（可以盤點、修改）、檢視者（只能看）。
 * - 每個人一組自己的連結碼（只存雜湊值）；有人離職，管理員在 App 按「移除權限」，只有他失效，其他人不用改；被移除的裝置下次連線會自動清除資料。
 * - 擁有者的手機不見了：Apps Script 左邊「專案設定」→ 最下面「指令碼屬性」→ 刪掉 SYNC_KEY，再從 App 重新開啟同步（其他人的權限會保留）。
 */
const RAW = '盤點紀錄'
const HEAD = ['盤點日期', '時間', '位置', '品名', '品牌', '型號', '尺寸／規格', '數量', '照片框數', '來源', '盤點ID']
const SYNC_ACTIONS = ['push', 'pull', 'members', 'invite', 'remove', 'setRole', 'rename', 'reissue']
const ROLE_NAME = { owner: '擁有者', manager: '管理員', editor: '編輯者', viewer: '檢視者' }

function doPost(e) {
  const data = JSON.parse(e.postData.contents)
  // 一次讀完全部設定（讀寫次數有每日上限）
  const props = PropertiesService.getScriptProperties()
  const P = props.getProperties()
  // 還沒開啟同步：只有使用者在 App 按「開啟同步」（hello）才能設定擁有者；背景同步不能設定，也讀不到資料
  if (!P.SYNC_KEY && data.action === 'hello' && data.key && String(data.key).length >= 16) {
    props.setProperty('SYNC_KEY', String(data.key))
    P.SYNC_KEY = String(data.key)
  }
  const syncOn = !!P.SYNC_KEY
  const who = syncOn ? whoIs(P, data.key) : null
  if (syncOn && !who) {
    const revoked = JSON.parse(P.SYNC_REVOKED || '[]').indexOf(sha(data.key || '')) >= 0
    return json({ ok: false, revoked: revoked, error: revoked ? '這台已經被移除權限，不能再同步' : '連結碼不對或已經失效：請找管理員要一個新的連結碼' })
  }
  if (!syncOn && (data.action === 'hello' || SYNC_ACTIONS.indexOf(data.action) >= 0)) return json({ ok: false, error: '還沒開啟同步：請在第一台裝置按「開啟多人同步」' })
  if (who) touchSeen(props, P, who)
  // 搬家（例如從個人帳號搬到公司帳號）：只有擁有者能做；搬完後，舊的位置會告訴每台 App 新網址，大家自動跟過去
  if (who && who.role === 'owner' && data.action === 'exportMembers') return json({ ok: true, members: membersOf(P), revokedHashes: JSON.parse(P.SYNC_REVOKED || '[]'), ownerName: P.SYNC_OWNER_NAME || '' })
  if (who && who.role === 'owner' && data.action === 'importMembers') {
    props.setProperty('SYNC_MEMBERS', JSON.stringify(data.members || []))
    props.setProperty('SYNC_REVOKED', JSON.stringify(data.revokedHashes || []))
    if (data.ownerName) props.setProperty('SYNC_OWNER_NAME', String(data.ownerName))
    return json({ ok: true })
  }
  if (who && who.role === 'owner' && data.action === 'moveTo') {
    if (data.url) props.setProperty('SYNC_MOVED', String(data.url))
    else props.deleteProperty('SYNC_MOVED')
    return json({ ok: true })
  }
  if (who && P.SYNC_MOVED) return json({ ok: false, moved: P.SYNC_MOVED, error: '資料已經搬到新的位置' })
  const me = who ? { role: who.role, name: who.name, id: who.id } : {}
  const canEdit = !who || who.role !== 'viewer'
  const canManage = who && (who.role === 'owner' || who.role === 'manager')
  if (data.action === 'hello') return json({ ok: true, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), me: me })
  if (data.action === 'pull') return json(Object.assign(syncPull(data, P), { me: me }))
  if (data.action === 'push') return json(canEdit ? Object.assign(syncPush(data, props, P), { me: me }) : { ok: false, viewer: true, me: me, error: '你是檢視者，只能看' })
  if (SYNC_ACTIONS.indexOf(data.action) >= 0) return json(canManage ? Object.assign(manage(data, props, P, who), { me: me }) : { ok: false, me: me, error: '只有擁有者和管理員可以管理共用的人' })
  if (!canEdit) return json({ ok: false, viewer: true, error: '你是檢視者，只能看' })
  const lock = LockService.getScriptLock()
  lock.waitLock(20000)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const sh = ensureRaw(ss)
    const rows = (data.rows || []).map((r) => {
      const row = r.slice(0, HEAD.length)
      row[0] = new Date(String(row[0]) + 'T00:00:00') // 盤點日期存成真的日期，才能排序、篩選
      return row
    })
    // 同一次盤點重新同步：先刪掉舊的那幾列
    const ids = {}
    rows.forEach((r) => (ids[String(r[10])] = true))
    const last = sh.getLastRow()
    if (last > 1) {
      const idCol = sh.getRange(2, 11, last - 1, 1).getValues()
      for (let i = idCol.length - 1; i >= 0; i--) if (ids[String(idCol[i][0])]) sh.deleteRow(i + 2)
    }
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.length).setValues(rows)
    sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).setNumberFormat('yyyy/mm/dd')
    ensureSummary(ss)
    if (data.items && data.items.length) writeItems(ss, data.items)
    return json({ ok: true, rows: rows.length, items: data.items ? data.items.length - 1 : 0 })
  } finally {
    lock.releaseLock()
  }
}

/** App 的「測試連線」會呼叫這個 */
function doGet() {
  return json({ ok: true, app: '拍照盤點', sheet: SpreadsheetApp.getActiveSpreadsheet().getName() })
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)
}

// ───────────── 共用的人與權限 ─────────────
/** 連結碼只存雜湊值（SHA-256）：就算有人看到設定，也拿不到別人的連結碼 */
function sha(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8)
    .map(function (b) {
      return ((b + 256) % 256).toString(16).padStart(2, '0')
    })
    .join('')
}
function membersOf(P) {
  return JSON.parse(P.SYNC_MEMBERS || '[]')
}
/** 這個連結碼是誰：擁有者（第一台）或共用名單裡的人 */
function whoIs(P, key) {
  if (!key) return null
  if (key === P.SYNC_KEY) return { id: 'owner', name: P.SYNC_OWNER_NAME || '擁有者', role: 'owner' }
  const h = sha(key)
  const m = membersOf(P).filter(function (x) {
    return x.h === h
  })[0]
  return m ? { id: m.id, name: m.name, role: m.role } : null
}
/** 最後上線時間：超過 10 分鐘才寫一次（省讀寫次數） */
function touchSeen(props, P, who) {
  const now = Date.now()
  if (who.id === 'owner') {
    if (now - Number(P.SYNC_OWNER_SEEN || 0) > 600000) props.setProperty('SYNC_OWNER_SEEN', String(now))
    return
  }
  const list = membersOf(P)
  const m = list.filter(function (x) {
    return x.id === who.id
  })[0]
  if (m && now - (m.seen || 0) > 600000) {
    m.seen = now
    props.setProperty('SYNC_MEMBERS', JSON.stringify(list))
  }
}
/** 管理共用的人（擁有者、管理員才能用） */
function manage(data, props, P, who) {
  const lock = LockService.getScriptLock()
  lock.waitLock(20000)
  try {
    // 鎖住後重新讀（別的管理員可能剛改過）
    const list = JSON.parse(props.getProperty('SYNC_MEMBERS') || '[]')
    const find = function (id) {
      return list.filter(function (x) {
        return x.id === id
      })[0]
    }
    const save = function () {
      props.setProperty('SYNC_MEMBERS', JSON.stringify(list))
    }
    const okRole = function (r) {
      return r === 'manager' || r === 'editor' || r === 'viewer'
    }
    if (data.action === 'members') {
      const people = [{ id: 'owner', name: P.SYNC_OWNER_NAME || '擁有者', role: 'owner', seen: Number(P.SYNC_OWNER_SEEN || 0) }].concat(
        list.map(function (m) {
          return { id: m.id, name: m.name, role: m.role, at: m.at, seen: m.seen || 0 }
        }),
      )
      return { ok: true, people: people }
    }
    if (data.action === 'invite') {
      const name = String(data.name || '').trim().slice(0, 30)
      if (!name) return { ok: false, error: '請填名字' }
      if (!okRole(data.role)) return { ok: false, error: '權限不對' }
      if (!data.newKey || String(data.newKey).length < 16) return { ok: false, error: '連結碼太短' }
      const id = Utilities.getUuid()
      list.push({ id: id, name: name, role: data.role, h: sha(data.newKey), at: Date.now(), seen: 0 })
      save()
      return { ok: true, id: id }
    }
    if (data.id === 'owner') {
      // 擁有者不能被移除、不能被改權限；只能改名字
      if (data.action === 'rename' && who.role === 'owner') {
        props.setProperty('SYNC_OWNER_NAME', String(data.name || '').trim().slice(0, 30) || '擁有者')
        return { ok: true }
      }
      return { ok: false, error: '擁有者不能被移除或改權限' }
    }
    const m = find(data.id)
    if (!m) return { ok: false, error: '找不到這個人（可能已經被移除）' }
    if (data.action === 'setRole') {
      if (!okRole(data.role)) return { ok: false, error: '權限不對' }
      m.role = data.role
    } else if (data.action === 'rename') {
      m.name = String(data.name || '').trim().slice(0, 30) || m.name
    } else if (data.action === 'reissue' || data.action === 'remove') {
      // 舊的連結碼列入黑名單：那台下次連線會自動清除資料
      const revoked = JSON.parse(props.getProperty('SYNC_REVOKED') || '[]')
      revoked.push(m.h)
      props.setProperty('SYNC_REVOKED', JSON.stringify(revoked.slice(-500)))
      if (data.action === 'remove') list.splice(list.indexOf(m), 1)
      else {
        if (!data.newKey || String(data.newKey).length < 16) return { ok: false, error: '連結碼太短' }
        m.h = sha(data.newKey)
        m.seen = 0
      }
    }
    save()
    return { ok: true }
  } finally {
    lock.releaseLock()
  }
}

// ───────────── 多台裝置同步：資料存在雲端硬碟，一筆一個檔案＋一個目錄檔（index.json） ─────────────
/**
 * 目錄：{ 鍵: { t: 裝置上的修改時間, s: 收到的順序（這邊的時間）, d: 哪一台送的, f: 檔案 ID, del: 是否已刪除 } }
 * 鍵例如 session:xxx（一次盤點，含照片）、item:xxx（品項）、sample:xxx（樣品照）、settings（儲位、品項清單）
 */
function syncFolder(P) {
  const props = PropertiesService.getScriptProperties()
  const id = P.SYNC_FOLDER
  if (id) {
    try {
      return DriveApp.getFolderById(id)
    } catch (e) {
      /* 資料夾被刪了：重新建立 */
    }
  }
  const folder = DriveApp.createFolder('拍照盤點同步資料（不要刪）')
  props.setProperty('SYNC_FOLDER', folder.getId())
  P.SYNC_FOLDER = folder.getId()
  return folder
}
function readIndex(folder) {
  const files = folder.getFilesByName('index.json')
  if (!files.hasNext()) return { file: null, map: {} }
  const file = files.next()
  return { file: file, map: JSON.parse(file.getBlob().getDataAsString() || '{}') }
}
function writeIndex(folder, idx) {
  const text = JSON.stringify(idx.map)
  if (idx.file) idx.file.setContent(text)
  else idx.file = folder.createFile(Utilities.newBlob(text, 'application/json', 'index.json'))
}

/** 上傳：比較新的才寫（以裝置上的修改時間為準）；舊檔案丟到垃圾桶 */
function syncPush(data, props, P) {
  const lock = LockService.getScriptLock()
  lock.waitLock(30000)
  try {
    const folder = syncFolder(P)
    const idx = readIndex(folder)
    // 順序號一定越來越大（同一毫秒兩次上傳也不會重複），下載時才不會漏；鎖住後重新讀一次，別人剛寫的才不會被蓋掉
    let seq = Math.max(Date.now(), Number(props.getProperty('SYNC_SEQ') || 0) + 1)
    let n = 0
    const skipped = []
    ;(data.records || []).forEach(function (r) {
      const cur = idx.map[r.k]
      if (cur && cur.t > r.t) {
        skipped.push(r.k) // 雲端已經有比較新的（別台改的）
        return
      }
      if (cur && cur.f) {
        try {
          DriveApp.getFileById(cur.f).setTrashed(true)
        } catch (e) {
          /* 檔案已經不在了 */
        }
      }
      let f = ''
      if (!r.del) f = folder.createFile(Utilities.newBlob(JSON.stringify(r.d), 'application/json', r.k.replace(/[^\w-]/g, '_') + '.json')).getId()
      idx.map[r.k] = { t: r.t, s: seq++, d: String(data.dev || ''), f: f, del: !!r.del }
      n++
    })
    writeIndex(folder, idx)
    props.setProperty('SYNC_SEQ', String(seq))
    return { ok: true, n: n, skipped: skipped }
  } finally {
    lock.releaseLock()
  }
}

/** 下載：別台送來、比 since 新的；一次最多約 8 MB 或 60 筆，more＝還有 */
function syncPull(data, P) {
  const folder = syncFolder(P)
  const idx = readIndex(folder)
  const since = Number(data.since) || 0
  const dev = String(data.dev || '')
  const all = Object.keys(idx.map)
    .map(function (k) {
      return { k: k, e: idx.map[k] }
    })
    .filter(function (x) {
      return x.e.s > since
    })
    .sort(function (a, b) {
      return a.e.s - b.e.s
    })
  const records = []
  let size = 0
  let next = since
  let i = 0
  for (; i < all.length; i++) {
    const x = all[i]
    if (records.length && (size > 8e6 || records.length >= 60)) break
    next = x.e.s
    if (x.e.d === dev) continue // 自己送的不用再下載
    let d = null
    if (!x.e.del && x.e.f) {
      try {
        const text = DriveApp.getFileById(x.e.f).getBlob().getDataAsString()
        size += text.length
        d = JSON.parse(text)
      } catch (e) {
        d = null
      }
    }
    records.push({ k: x.k, t: x.e.t, del: !!x.e.del, d: d })
  }
  return { ok: true, records: records, next: next, more: i < all.length }
}

function ensureRaw(ss) {
  const sh = ss.getSheetByName(RAW) || ss.insertSheet(RAW, 0)
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD)
    sh.setFrozenRows(1)
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#e8f0fe')
  }
  return sh
}

/** 品項庫：整張覆蓋（第一列是標題）；盤虧標紅、該叫貨標橘 */
function writeItems(ss, table) {
  const sh = ss.getSheetByName('品項庫') || ss.insertSheet('品項庫', 1)
  sh.clear()
  const width = table[0].length
  const rows = table.map((r) => {
    const row = r.slice(0, width)
    while (row.length < width) row.push('')
    return row
  })
  sh.getRange(1, 1, rows.length, width).setValues(rows)
  sh.setFrozenRows(1)
  sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#e8f0fe')
  const status = table[0].indexOf('狀態')
  if (rows.length < 2 || status < 0) return
  // 一次設定整張的底色（一列一列設會很慢）
  const colors = rows.slice(1).map((r) => {
    const s = String(r[status] || '')
    const c = s.indexOf('盤虧') >= 0 ? '#fde8e8' : s.indexOf('該叫貨') >= 0 ? '#fff4e0' : null
    return r.map(() => c)
  })
  sh.getRange(2, 1, colors.length, width).setBackgrounds(colors)
}

function ensureSummary(ss) {
  if (!ss.getSheetByName('總表')) {
    const s = ss.insertSheet('總表')
    s.getRange('A1').setValue('每個品項在每次盤點日期各幾件（自動計算，不用改）').setFontWeight('bold')
    s.getRange('A3').setFormula('=QUERY(\'盤點紀錄\'!A:K,"select D, G, sum(H) where D is not null group by D, G pivot A",1)')
    s.setFrozenRows(3)
  }
  if (!ss.getSheetByName('最新一次')) {
    const s = ss.insertSheet('最新一次')
    s.getRange('A1').setValue('最近一次盤點日期').setFontWeight('bold')
    s.getRange('B1').setFormula("=MAX('盤點紀錄'!A2:A)").setNumberFormat('yyyy/mm/dd')
    s.getRange('A3').setFormula(
      '=QUERY(\'盤點紀錄\'!A:K,"select D, G, E, F, sum(H) where A = date \'"&TEXT(B1,"yyyy-mm-dd")&"\' group by D, G, E, F order by sum(H) desc label D \'品名\', G \'尺寸／規格\', E \'品牌\', F \'型號\', sum(H) \'數量\'",1)',
    )
    s.setFrozenRows(3)
  }
}
