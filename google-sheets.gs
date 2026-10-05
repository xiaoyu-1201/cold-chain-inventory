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
 * 會自動建立五個工作表（每次同步自動更新；手動改會先跳警告）：
 * - 總覽：數字卡片（品項、件數、該叫貨、盤虧、盤盈、儲位已盤）、該叫貨／盤差最大前 10、儲位盤點進度、圖表
 * - 庫存：每個品項的實盤、帳面、差異、叫貨點、在哪裡、最近盤點、盤點人（盤虧紅、該叫貨橘）
 * - 儲位庫存：每一格放了什麼、幾個、誰什麼時候盤的
 * - 盤差報告：實盤≠帳面的品項；後三欄「原因說明、處理方式、主管確認」給主管填，重新同步會保留
 * - 盤點紀錄：每一次、每一種商品一列（原始資料，不要手動改欄位順序）
 *
 * 多人同步（手機、電腦、同事看到同一份資料）：
 * - 盤點紀錄（含照片）、品項庫、儲位、樣品照存在你自己的 Google 雲端硬碟「拍照盤點同步資料（不要刪）」資料夾。
 * - 權限像 Google 雲端硬碟的「共用」：擁有者（第一台）、管理員（可以邀請／移除人）、編輯者（可以盤點、修改）、檢視者（只能看）。
 * - 每個人一組自己的連結碼（只存雜湊值）；有人離職，管理員在 App 按「移除權限」，只有他失效，其他人不用改；被移除的裝置下次連線會自動清除資料。
 * - 擁有者的手機不見了：Apps Script 左邊「專案設定」→ 最下面「指令碼屬性」→ 刪掉 SYNC_KEY，再從 App 重新開啟同步（其他人的權限會保留）。
 */
const RAW = '盤點紀錄'
const HEAD = ['盤點日期', '時間', '位置', '品名', '品牌', '型號', '尺寸／規格', '數量', '照片框數', '來源', '盤點ID', '盤點人']
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
  // 名單（只有名字，給每台 App 顯示「誰盤的」、選盤點人；名字只有擁有者／管理員能改）
  const roster = who
    ? [{ id: 'owner', name: P.SYNC_OWNER_NAME || '擁有者' }].concat(
        membersOf(P).map(function (m) {
          return { id: m.id, name: m.name }
        }),
      )
    : []
  if (data.action === 'hello') return json({ ok: true, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), me: me, roster: roster })
  if (data.action === 'pull') return json(Object.assign(syncPull(data, P), { me: me, roster: roster }))
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
      while (row.length < HEAD.length) row.push('') // 舊版 App 少送「盤點人」
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
    if (data.report && data.report.items) writeReports(ss, data.report)
    else if (data.items && data.items.length) writeTable(ss, '庫存', 1, data.items, { color: COLOR.green }) // 舊版 App
    return json({ ok: true, rows: rows.length, items: data.report ? data.report.items.length - 1 : 0 })
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
  const sh = ss.getSheetByName(RAW) || ss.insertSheet(RAW, ss.getSheets().length) // 原始資料放最後面
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD)
    sh.setFrozenRows(1)
  }
  // 以前建立的工作表少了新欄位（例如「盤點人」）：補上標題
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#e8f0fe')
  return sh
}

// ───────────── 公司用的報表：總覽、庫存、儲位庫存、盤差報告 ─────────────
// App 每次同步送來三張表（第一列是標題）：items（庫存）、stock（儲位庫存）、locs（儲位進度）。
// 時間欄位送 {$t: 毫秒}，這邊轉成真的日期（才能排序、篩選）。
const COLOR = { blue: '#0a84ff', green: '#30a46c', orange: '#f76b15', red: '#e5484d', gray: '#8e8e93', head: '#e8f0fe', card: '#f2f2f7', loss: '#fde8e8', gain: '#e8f0fe', order: '#fff4e0' }

function cellOf(v) {
  return v && typeof v === 'object' && v.$t ? new Date(v.$t) : v
}
function rowsOf(table) {
  const width = table[0].length
  return table.map(function (r) {
    const row = r.slice(0, width).map(cellOf)
    while (row.length < width) row.push('')
    return row
  })
}
function styleHeader(sh, width) {
  sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground(COLOR.head).setVerticalAlignment('middle')
  sh.setFrozenRows(1)
}
/** 自動產生的工作表：有人手動改會先跳警告（不會擋住，只是提醒） */
function protectWarn(sh, unprotectedA1) {
  if (sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) return
  const p = sh.protect().setDescription('由拍照盤點 App 自動產生：請在 App 裡改').setWarningOnly(true)
  if (unprotectedA1) p.setUnprotectedRanges([sh.getRange(unprotectedA1)])
}
function sheetAt(ss, name, index) {
  return ss.getSheetByName(name) || ss.insertSheet(name, Math.min(index, ss.getSheets().length))
}
/** 整張重寫一張表：標題、凍結、篩選、時間格式、欄寬 */
function writeTable(ss, name, index, table, opt) {
  const sh = sheetAt(ss, name, index)
  if (sh.getFilter()) sh.getFilter().remove()
  sh.clear()
  const rows = rowsOf(table)
  const width = rows[0].length
  sh.getRange(1, 1, rows.length, width).setValues(rows)
  styleHeader(sh, width)
  ;(opt.timeCols || []).forEach(function (h) {
    const c = rows[0].indexOf(h)
    if (c >= 0 && rows.length > 1) sh.getRange(2, c + 1, rows.length - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm')
  })
  if (rows.length > 1) sh.getRange(1, 1, rows.length, width).createFilter()
  sh.setTabColor(opt.color)
  sh.autoResizeColumns(1, width)
  protectWarn(sh)
  return { sh: sh, rows: rows }
}

function writeReports(ss, rep) {
  // 舊版的「品項庫」改名成「庫存」
  const old = ss.getSheetByName('品項庫')
  if (old && !ss.getSheetByName('庫存')) old.setName('庫存')
  const inv = writeTable(ss, '庫存', 1, rep.items, { timeCols: ['最近盤點'], color: COLOR.green })
  // 庫存：盤虧標紅、該叫貨標橘（一次設定整張底色，一列一列設會很慢）
  const st = inv.rows[0].indexOf('狀態')
  if (inv.rows.length > 1 && st >= 0) {
    inv.sh.getRange(2, 1, inv.rows.length - 1, inv.rows[0].length).setBackgrounds(
      inv.rows.slice(1).map(function (r) {
        const s = String(r[st] || '')
        const c = s.indexOf('盤虧') >= 0 ? COLOR.loss : s.indexOf('該叫貨') >= 0 ? COLOR.order : null
        return r.map(function () {
          return c
        })
      }),
    )
  }
  writeTable(ss, '儲位庫存', 2, rep.stock, { timeCols: ['盤點時間'], color: COLOR.blue })
  writeDiff(ss, rep.items)
  writeDashboard(ss, rep)
  const raw = ss.getSheetByName(RAW)
  if (raw) raw.setTabColor(COLOR.gray)
  // 新試算表預設的空白工作表（工作表1／Sheet1）沒用到就刪掉
  ss.getSheets().forEach(function (s) {
    if (/^(工作表|Sheet)\d+$/.test(s.getName()) && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s)
  })
}

/** 盤差報告：只列實盤≠帳面，差越多排越前面；後面三欄給主管填（重新同步也會保留） */
function writeDiff(ss, items) {
  const H = items[0]
  const col = function (h) {
    return H.indexOf(h)
  }
  const BASE = ['料號', '品名', '品牌', '型號', '尺寸／規格', '實盤', '帳面', '差異', '差異比例', '在哪裡（位置 數量）', '最近盤點', '盤點人']
  const NOTE = ['原因說明（請填）', '處理方式（請填）', '主管確認']
  const sh = sheetAt(ss, '盤差報告', 3)
  // 先把主管填過的留下來（用料號對）
  const notes = {}
  const last = sh.getLastRow()
  if (last > 1) {
    sh.getRange(2, 1, last - 1, BASE.length + NOTE.length)
      .getValues()
      .forEach(function (r) {
        if (r[0]) notes[String(r[0])] = r.slice(BASE.length)
      })
  }
  if (sh.getFilter()) sh.getFilter().remove()
  sh.clear()
  const diffs = items
    .slice(1)
    .filter(function (r) {
      return r[col('差異')] !== '' && Number(r[col('差異')]) !== 0
    })
    .sort(function (a, b) {
      return Math.abs(b[col('差異')]) - Math.abs(a[col('差異')])
    })
  const rows = [BASE.concat(NOTE)].concat(
    diffs.map(function (r) {
      const d = Number(r[col('差異')])
      const book = Number(r[col('帳面')])
      return [r[col('料號')], r[col('品名')], r[col('品牌')], r[col('型號')], r[col('尺寸／規格')], r[col('實盤')], r[col('帳面')], d, book ? d / book : '', r[col('在哪裡（位置 數量）')], cellOf(r[col('最近盤點')]), r[col('盤點人')]].concat(notes[String(r[col('料號')])] || ['', '', false])
    }),
  )
  const width = rows[0].length
  sh.getRange(1, 1, rows.length, width).setValues(rows)
  styleHeader(sh, width)
  sh.getRange(1, BASE.length + 1, 1, NOTE.length).setBackground(COLOR.order)
  if (rows.length > 1) {
    const n = rows.length - 1
    sh.getRange(2, 9, n, 1).setNumberFormat('0%')
    sh.getRange(2, 11, n, 1).setNumberFormat('yyyy/mm/dd hh:mm')
    sh.getRange(2, width, n, 1).insertCheckboxes()
    sh.getRange(2, 1, n, BASE.length).setBackgrounds(
      diffs.map(function (r) {
        const c = Number(r[col('差異')]) < 0 ? COLOR.loss : COLOR.gain
        return BASE.map(function () {
          return c
        })
      }),
    )
    sh.getRange(1, 1, rows.length, width).createFilter()
  } else {
    sh.getRange(2, 1).setValue('目前沒有盤差：實盤跟帳面都一樣（或還沒設定帳面數）。').setFontColor(COLOR.gray)
  }
  sh.setTabColor(COLOR.red)
  sh.autoResizeColumns(1, width)
  protectWarn(sh, 'M:O')
}

/** 總覽：給主管一眼看懂（數字卡片、需要處理的事、儲位盤點進度、圖表） */
function writeDashboard(ss, rep) {
  const sh = sheetAt(ss, '總覽', 0)
  sh.getCharts().forEach(function (c) {
    sh.removeChart(c)
  })
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart()
  sh.clear()
  sh.setHiddenGridlines(true)
  sh.setTabColor(COLOR.blue)
  const H = rep.items[0]
  const col = function (h) {
    return H.indexOf(h)
  }
  const items = rep.items.slice(1)
  const num = function (v) {
    return Number(v) || 0
  }
  const sum = function (list, f) {
    return list.reduce(function (n, r) {
      return n + f(r)
    }, 0)
  }
  const losses = items.filter(function (r) {
    return r[col('差異')] !== '' && num(r[col('差異')]) < 0
  })
  const gains = items.filter(function (r) {
    return r[col('差異')] !== '' && num(r[col('差異')]) > 0
  })
  const orders = items.filter(function (r) {
    return String(r[col('狀態')]).indexOf('該叫貨') >= 0
  })
  const locs = rep.locs.slice(1)
  const counted = locs.filter(function (r) {
    return r[6] !== '還沒盤'
  })
  const stale = locs.filter(function (r) {
    return r[6] === '超過 30 天沒盤'
  })
  // 欄寬：A 留白，B～M 六張卡片（每張兩欄）
  sh.setColumnWidth(1, 16)
  sh.setColumnWidths(2, 12, 92)
  sh.getRange('B1').setValue('盤點總覽').setFontSize(22).setFontWeight('bold')
  sh.getRange('B2').setValue('最後同步：').setFontColor(COLOR.gray)
  sh.getRange('C2').setValue(new Date(rep.at || Date.now())).setNumberFormat('yyyy/mm/dd hh:mm').setFontColor(COLOR.gray).setHorizontalAlignment('left')
  sh.getRange('E2').setValue('這一頁由拍照盤點 App 自動產生，不用改；細節看後面的工作表。').setFontColor(COLOR.gray)
  const cards = [
    ['品項', items.length, '種商品', COLOR.blue],
    ['實盤總件數', sum(items, function (r) {
      return num(r[col('實盤')])
    }), '件', COLOR.blue],
    ['該叫貨', orders.length, '項（看「庫存」橘色）', orders.length ? COLOR.orange : COLOR.green],
    ['盤虧', losses.length, '項，少 ' + -sum(losses, function (r) {
      return num(r[col('差異')])
    }) + ' 件', losses.length ? COLOR.red : COLOR.green],
    ['盤盈', gains.length, '項，多 ' + sum(gains, function (r) {
      return num(r[col('差異')])
    }) + ' 件', gains.length ? COLOR.blue : COLOR.green],
    ['儲位已盤', counted.length + '／' + locs.length, stale.length ? stale.length + ' 格超過 30 天沒盤' : '格', stale.length ? COLOR.orange : COLOR.green],
  ]
  cards.forEach(function (c, i) {
    const colNo = 2 + i * 2
    sh.getRange(4, colNo, 3, 2).setBackground(COLOR.card)
    sh.getRange(4, colNo, 1, 2).merge().setValue(c[0]).setFontColor(COLOR.gray).setFontSize(11)
    sh.getRange(5, colNo, 1, 2).merge().setValue(c[1]).setFontSize(26).setFontWeight('bold').setFontColor(c[3]).setHorizontalAlignment('left')
    sh.getRange(6, colNo, 1, 2).merge().setValue(c[2]).setFontColor(COLOR.gray).setFontSize(10)
  })
  // 需要處理：左邊該叫貨、右邊盤差最大
  const section = function (row, colNo, title, head, data, empty) {
    sh.getRange(row, colNo).setValue(title).setFontWeight('bold').setFontSize(13)
    sh.getRange(row + 1, colNo, 1, head.length).setValues([head]).setFontWeight('bold').setBackground(COLOR.head)
    if (data.length) sh.getRange(row + 2, colNo, data.length, head.length).setValues(data)
    else sh.getRange(row + 2, colNo).setValue(empty).setFontColor(COLOR.gray)
  }
  section(
    8,
    2,
    '該叫貨（前 10 項）',
    ['料號', '品名', '規格', '現在', '叫貨點'],
    orders.slice(0, 10).map(function (r) {
      return [r[col('料號')], r[col('品名')], r[col('尺寸／規格')], r[col('帳面')] !== '' ? r[col('帳面')] : r[col('實盤')], r[col('剩幾個要叫貨（安全庫存）')]]
    }),
    '沒有要叫貨的',
  )
  section(
    8,
    8,
    '盤差最大（前 10 項）',
    ['料號', '品名', '實盤', '帳面', '差異'],
    losses
      .concat(gains)
      .sort(function (a, b) {
        return Math.abs(b[col('差異')]) - Math.abs(a[col('差異')])
      })
      .slice(0, 10)
      .map(function (r) {
        return [r[col('料號')], r[col('品名')], r[col('實盤')], r[col('帳面')], r[col('差異')]]
      }),
    '沒有盤差',
  )
  // 儲位盤點進度
  const locRow = 22
  sh.getRange(locRow, 2).setValue('儲位盤點進度').setFontWeight('bold').setFontSize(13)
  const locTable = rowsOf(rep.locs)
  sh.getRange(locRow + 1, 2, locTable.length, locTable[0].length).setValues(locTable)
  sh.getRange(locRow + 1, 2, 1, locTable[0].length).setFontWeight('bold').setBackground(COLOR.head)
  if (locTable.length > 1) {
    sh.getRange(locRow + 2, 6, locTable.length - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm')
    sh.getRange(locRow + 2, 8, locTable.length - 1, 1).setFontColors(
      locTable.slice(1).map(function (r) {
        return [r[6] === '已盤' ? COLOR.green : r[6] === '還沒盤' ? COLOR.gray : COLOR.orange]
      }),
    )
  } else sh.getRange(locRow + 2, 2).setValue('還沒建立儲位：在 App「品項 → 儲位」建立').setFontColor(COLOR.gray)
  // 圖表：件數最多的品名（資料放在最右邊 Y:Z）
  const byLabel = {}
  items.forEach(function (r) {
    byLabel[r[col('品名')]] = (byLabel[r[col('品名')]] || 0) + num(r[col('實盤')])
  })
  const top = Object.keys(byLabel)
    .map(function (k) {
      return [k, byLabel[k]]
    })
    .sort(function (a, b) {
      return b[1] - a[1]
    })
    .slice(0, 10)
  if (top.length) {
    sh.getRange(1, 25, top.length + 1, 2).setValues([['品名', '實盤件數']].concat(top)).setFontColor(COLOR.gray)
    sh.insertChart(
      sh
        .newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(sh.getRange(1, 25, top.length + 1, 2))
        .setPosition(locRow, 10, 0, 0)
        .setOption('title', '件數最多的品名（前 10）')
        .setOption('legend', { position: 'none' })
        .setOption('colors', [COLOR.blue])
        .build(),
    )
  }
  protectWarn(sh)
}
