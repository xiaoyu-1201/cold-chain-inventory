/**
 * 聖佳智慧庫存（原名拍照盤點）→ Google 試算表
 * 用法（只要做一次）：
 * 1. 開一個新的 Google 試算表 → 上方「擴充功能」→「Apps Script」
 * 2. 把這整段貼上（取代原本的內容）→ 存檔
 * 3. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    執行身分：我　／　誰可以存取：所有人 → 部署 → 第一次會要你授權（選自己的帳號 → 進階 → 前往）
 * 4. 複製「網頁應用程式網址」（https://script.google.com/macros/s/…/exec），貼到聖佳智慧庫存 App 的「設定 → Google 試算表」
 * 之後每次在 App 按「完成」，這次的盤點就會寫進來；同一次盤點重複送，會先刪掉舊的再寫，不會重複。
 *
 * 會自動建立五個工作表（每次同步自動更新；手動改會先跳警告）：
 * - 總覽：數字卡片（品項、件數、該叫貨、盤虧、盤盈、儲位已盤）、該叫貨／盤差最大前 10、儲位盤點進度、圖表
 * - 庫存：每個品項的實盤、帳面、差異、叫貨點、在哪裡、最近盤點、盤點人（盤虧紅、該叫貨橘）
 * - 儲位庫存：每一格放了什麼、幾個、誰什麼時候盤的
 * - 盤差報告：實盤≠帳面的品項＋複盤狀態；後三欄「原因說明、處理方式、主管確認」給主管填，重新同步會保留（原因沒填時，帶入 App 複盤時選的原因）
 * - 盤點紀錄：每一次、每一種商品一列（原始資料，不要手動改欄位順序）
 *
 * 多人同步（手機、電腦、同事看到同一份資料）：
 * - 盤點紀錄（含照片）、品項庫、儲位、樣品照存在你自己的 Google 雲端硬碟「拍照盤點同步資料（不要刪）」資料夾。
 * - 權限像 Google 雲端硬碟的「共用」：擁有者（第一台）、管理員（可以邀請／移除人）、編輯者（可以盤點、修改）、檢視者（只能看）。
 * - 每個人一組自己的連結碼（只存雜湊值）；有人離職，管理員在 App 按「移除權限」，只有他失效，其他人不用改；被移除的裝置下次連線會自動清除資料。
 * - 擁有者的手機不見了：Apps Script 左邊「專案設定」→ 最下面「指令碼屬性」→ 刪掉 SYNC_KEY，再從 App 重新開啟同步（其他人的權限會保留）。
 *
 * 共用 AI 金鑰（4.5 起）：擁有者在 App「設定 → 把金鑰放到雲端」→ 金鑰存在這裡的指令碼屬性 GEMINI_KEY。
 * 同事拍照時，請求送到這裡，由這裡拿金鑰去問 Gemini，同事的手機不會有金鑰；移除權限的人就不能再用。
 * 第一次更新到這一版，部署時 Google 會再問一次授權（多了「連到外部網站」，用來連 Gemini），按允許即可。
 *
 * 4.6.1：共用金鑰加上費用上限（只准 Flash／Flash-Lite／Pro、每人每天 300 次、一次最多 8 MB）；
 * 編輯者不能改盲盤／複盤規則、不能匯入正航產品表、不能一次刪光盤點紀錄（雲端也擋）。
 * 更新方法：整段重新貼上 → 存檔 →「部署」→「管理部署作業」→ 鉛筆 → 版本選「新版本」→ 部署（網址不會變）。
 *
 * 4.7.6：同步變快——很多筆資料、照片「同時」讀寫（Drive API），不再一筆一筆來；
 * 記住目錄檔的位置，不用每次搜尋。Drive API 用不了會自動改回舊方法（結果一樣，只是比較慢），
 * App 的設定頁會提醒擁有者：左邊「服務」→ 加入「Drive API」→ 再部署一次新版本，就會打開。
 * 第一次更新到這一版，Google 可能會再問一次授權，按允許即可。
 */
const RAW = '盤點紀錄'
const HEAD = ['盤點日期', '時間', '位置', '品名', '品牌', '型號', '尺寸／規格', '數量', '照片框數', '來源', '盤點ID', '盤點人']
const SYNC_ACTIONS = ['push', 'pull', 'photos', 'members', 'invite', 'remove', 'setRole', 'rename', 'reissue']
const SYNC_VER = 3 // 同步格式版本：2＝照片分開存；3＝同時讀寫（變快，格式跟 2 一樣）
const isPhotoKey = function (k) {
  return String(k).indexOf('photo:') === 0
}
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
  // 共用 AI 金鑰：ai＝擁有者有沒有放；每個回覆都帶，App 才知道能不能用
  if (!who && (data.action === 'ai' || data.action === 'aiKey')) return json({ ok: false, error: '還沒開啟同步，不能用共用的 AI 金鑰' })
  const ai = !!P.GEMINI_KEY
  if (who && data.action === 'aiKey') {
    if (who.role !== 'owner') return json({ ok: false, ai: ai, me: me, error: '只有擁有者可以設定共用的 AI 金鑰' })
    if (data.aiKey) props.setProperty('GEMINI_KEY', String(data.aiKey))
    else props.deleteProperty('GEMINI_KEY')
    return json({ ok: true, ai: !!data.aiKey, me: me })
  }
  if (who && data.action === 'ai') return json(canEdit ? Object.assign(aiProxy(data, P, props, who), { ai: ai, me: me }) : { ok: false, viewer: true, ai: ai, me: me, error: '你是檢視者，只能看' })
  // ver 2：盤點的照片分開存（photo:盤點ID:照片ID），下載只拿資料，照片用 photos 另外要；App 看到 ver 才改用新格式上傳
  // fast：有沒有用到「同時讀寫」（4.7.6）；沒有的話，App 會教擁有者打開 Drive API
  if (data.action === 'hello') return json({ ok: true, ver: SYNC_VER, ai: ai, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), me: me, roster: roster, fast: driveApiEnabled() })
  if (data.action === 'pull') return json(Object.assign(syncPull(data, P), { ver: SYNC_VER, ai: ai, me: me, roster: roster, fast: driveApiEnabled() }))
  if (data.action === 'photos') return json(Object.assign(syncPhotos(data, P), { ver: SYNC_VER, ai: ai, me: me }))
  if (data.action === 'push') return json(canEdit ? Object.assign(syncPush(data, props, P, !!canManage), { ver: SYNC_VER, ai: ai, me: me }) : { ok: false, viewer: true, ai: ai, me: me, error: '你是檢視者，只能看' })
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

/**
 * 共用 AI 金鑰：幫 App 問 Gemini（金鑰只在這裡）。只准兩種請求：列出模型、看圖回答（generateContent），
 * 不能拿這把金鑰做別的事。費用也有上限（擁有者付錢）：
 * - 模型只准 Gemini 的 Flash／Flash-Lite／Pro（不准畫圖、語音那些比較貴的）
 * - 不准用工具（Google 搜尋等）、快取；一次最多回 8192 個字詞、只回一個答案
 * - 一次最多 8 MB（一張照片＋30 張樣品照大約 2～4 MB）
 * - 每個人每天最多 300 次（可以在「專案設定 → 指令碼屬性」加 AI_DAILY_LIMIT 改次數）
 */
const AI_MAX_BYTES = 8 * 1024 * 1024
const AI_MAX_OUTPUT = 8192
const AI_DAILY_LIMIT = 300
function aiModelOk(name) {
  const m = String(name || '').replace(/^models\//, '')
  return /^gemini-(\d+(\.\d+)?-)?(flash|flash-lite|pro)(-[a-z0-9.\-]+)?$/.test(m) && !/image|tts|audio|live|embedding|thinking|8b|computer|robotics/.test(m)
}
function aiProxy(data, P, props, who) {
  if (!P.GEMINI_KEY) return { ok: false, status: 403, error: '擁有者還沒把 AI 金鑰放到雲端（或已經停止共用）' }
  const path = typeof data.path === 'string' ? data.path : ''
  const list = /^models\?pageSize=\d+$/.test(path)
  const gen = /^models\/([\w.\-]+):generateContent$/.exec(path)
  if (!list && !gen) return { ok: false, status: 400, error: '不支援的 AI 請求' }
  if (gen && !aiModelOk(gen[1])) return { ok: false, status: 404, error: '共用金鑰不能用這個模型（' + gen[1] + '）：只能用 Gemini 的 Flash、Flash-Lite、Pro' }
  const opt = { method: 'get', headers: { 'x-goog-api-key': P.GEMINI_KEY }, muteHttpExceptions: true }
  if (gen) {
    const raw = typeof data.body === 'string' ? data.body : JSON.stringify(data.body || {})
    if (raw.length > AI_MAX_BYTES) return { ok: false, status: 400, error: '這次送的照片太大（超過 8 MB）：請少放一些樣品照再試' }
    let body
    try {
      body = JSON.parse(raw)
    } catch (e) {
      body = null
    }
    if (!body || typeof body !== 'object' || !Array.isArray(body.contents)) return { ok: false, status: 400, error: 'AI 請求的格式不對' }
    delete body.tools
    delete body.toolConfig
    delete body.cachedContent
    const gc = body.generationConfig && typeof body.generationConfig === 'object' ? body.generationConfig : {}
    gc.maxOutputTokens = Math.min(Number(gc.maxOutputTokens) || AI_MAX_OUTPUT, AI_MAX_OUTPUT)
    gc.candidateCount = 1
    body.generationConfig = gc
    // 每人每天的次數：先算再問（鎖住只為了加一，不會卡住別人）
    const used = aiCount(props, P, who)
    if (used < 0) return { ok: false, status: 429, quota: true, error: '今天用共用 AI 的次數已經到上限（每人每天 ' + aiLimit(P) + ' 次），明天會自動恢復；急用請找擁有者' }
    opt.method = 'post'
    opt.contentType = 'application/json'
    opt.payload = JSON.stringify(body)
  }
  let res
  try {
    res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/' + path, opt)
  } catch (e) {
    // 連不到 Gemini（逾時、Google 這邊的每日上限）：照 Gemini 忙線的格式回，App 會等一下再試，不會顯示看不懂的錯誤頁
    return { ok: true, status: 503, body: JSON.stringify({ error: { code: 503, status: 'UNAVAILABLE', message: 'Apps Script 連不到 Gemini：' + String((e && e.message) || e).slice(0, 150) } }) }
  }
  const status = res.getResponseCode()
  let text = res.getContentText()
  // 列模型：只給 App 能用的那些（App 就不會挑到共用金鑰不能用的模型）
  if (list && status === 200) {
    try {
      const all = JSON.parse(text)
      all.models = (all.models || []).filter(function (m) {
        return aiModelOk(m.name)
      })
      text = JSON.stringify(all)
    } catch (e) {
      /* 回覆不是 JSON：原樣給 App */
    }
  }
  return { ok: true, status: status, body: text }
}
function aiLimit(P) {
  return Number(P.AI_DAILY_LIMIT) > 0 ? Number(P.AI_DAILY_LIMIT) : AI_DAILY_LIMIT
}
/** 這個人今天第幾次用共用 AI；超過上限回 -1（台灣時間每天 0 點重算） */
function aiCount(props, P, who) {
  const today = new Date(Date.now() + 8 * 3600 * 1000).toISOString().slice(0, 10)
  // 用試算表的鎖（跟同步用的鎖分開）：同步上傳照片要鎖比較久，問 AI 不用跟著等
  const lock = (LockService.getDocumentLock && LockService.getDocumentLock()) || LockService.getScriptLock()
  if (!lock.tryLock(10000)) return 0 // 一直等不到（很少見）：這次先不算，不要讓人卡住
  try {
    let usage = {}
    try {
      usage = JSON.parse(props.getProperty('AI_USAGE') || '{}') || {}
    } catch (e) {
      usage = {}
    }
    if (usage.day !== today) usage = { day: today, n: {} }
    const id = String(who.id)
    const n = (usage.n[id] || 0) + 1
    if (n > aiLimit(P)) return -1
    usage.n[id] = n
    props.setProperty('AI_USAGE', JSON.stringify(usage))
    return n
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
  const old = membersOf(P).filter(function (x) {
    return x.id === who.id
  })[0]
  if (!old || now - (old.seen || 0) <= 600000) return
  // 寫回整份名單之前先鎖住、重新讀一次：不然剛好有管理員在移除別人，會把被移除的人寫回去
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(3000)) return // 等不到就算了，下次再記
  try {
    const list = JSON.parse(props.getProperty('SYNC_MEMBERS') || '[]')
    const m = list.filter(function (x) {
      return x.id === who.id
    })[0]
    if (!m) return
    m.seen = now
    props.setProperty('SYNC_MEMBERS', JSON.stringify(list))
  } finally {
    lock.releaseLock()
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
/** Drive 說「找不到這個 ID」（真的被刪了）；其他錯誤（忙線、呼叫太多次）都當成一時出錯 */
const NOT_FOUND = /No item with the given ID|not found|找不到/i
function syncFolder(P) {
  const props = PropertiesService.getScriptProperties()
  const id = P.SYNC_FOLDER
  if (id) {
    // Drive 一時出錯（忙線、呼叫太多次）不能當成「資料夾被刪了」：不然會另建一個空資料夾，
    // 雲端的資料看起來全部不見（4.7.6 code review）。多試兩次；確定是「找不到」才重建，其他錯誤整個請求失敗、App 等一下再試
    let err = null
    for (let i = 0; i < 3; i++) {
      try {
        return DriveApp.getFolderById(id)
      } catch (e) {
        err = e
        if (i < 2) Utilities.sleep(500)
      }
    }
    if (!NOT_FOUND.test(String((err && err.message) || err))) throw err
  }
  const folder = DriveApp.createFolder('拍照盤點同步資料（不要刪）')
  props.setProperty('SYNC_FOLDER', folder.getId())
  P.SYNC_FOLDER = folder.getId()
  return folder
}
/**
 * 讀目錄檔（4.7.6 加快）：記住 index.json 的檔案 ID，直接打開，
 * 不用每次先開資料夾、再用檔名搜尋（每次同步省 1 秒左右）；打不開才用檔名找。
 * 還沒同步過（沒有資料夾）：回空的，不先建資料夾（建資料夾留給上傳，在鎖裡面做）。
 */
function readIndex(P) {
  if (P.SYNC_INDEX) {
    let f = null
    try {
      f = DriveApp.getFileById(P.SYNC_INDEX)
    } catch (e) {
      // 確定被刪了才往下用檔名找；Drive 一時出錯：整個請求失敗、App 等一下再試
      // （不能當成空的：上傳會蓋出一份只有新資料的目錄，大家的資料看起來就不見了）
      if (!NOT_FOUND.test(String((e && e.message) || e))) throw e
    }
    if (f) return { file: f, map: JSON.parse(f.getBlob().getDataAsString() || '{}') }
  }
  if (!P.SYNC_FOLDER) return { file: null, map: {} }
  const files = syncFolder(P).getFilesByName('index.json')
  if (!files.hasNext()) return { file: null, map: {} }
  const file = files.next()
  PropertiesService.getScriptProperties().setProperty('SYNC_INDEX', file.getId())
  P.SYNC_INDEX = file.getId()
  return { file: file, map: JSON.parse(file.getBlob().getDataAsString() || '{}') }
}
function writeIndex(P, idx) {
  const text = JSON.stringify(idx.map)
  if (idx.file) idx.file.setContent(text)
  else idx.file = syncFolder(P).createFile(Utilities.newBlob(text, 'application/json', 'index.json'))
  const id = idx.file.getId()
  if (P.SYNC_INDEX !== id) {
    PropertiesService.getScriptProperties().setProperty('SYNC_INDEX', id)
    P.SYNC_INDEX = id
  }
}

// ───────────── 一次同時讀寫很多個檔案（4.7.6：同步變快的主要原因） ─────────────
/**
 * 以前每一筆資料、每一張照片都要一個一個跟雲端硬碟要（一筆約 0.3～1 秒），
 * 現在用 Drive API 同時送出去（UrlFetchApp.fetchAll），幾十筆也只要一兩秒。
 * - Drive API 用不了（例如 Google 那邊沒開）：自動改回一個一個來，結果一樣、只是比較慢；記 6 小時，不會每次都先試。
 *   擁有者可以在 Apps Script 左邊「服務」加入「Drive API」打開它。
 * - 每讀、建、丟一個檔算一次「連到外部網站」（一般帳號每天 2 萬次，跟共用 AI 金鑰共用）：
 *   同時讀寫每天最多用 1 萬次，超過就改回一個一個來，剩下的留給 AI。
 */
const DRIVE_API = 'https://www.googleapis.com/drive/v3/files'
const DRIVE_UPLOAD = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id'
const FETCH_CHUNK = 20 // 一次同時送幾個
const UPLOAD_GROUP_BYTES = 4e6 // 一次同時上傳最多約 4 MB（照片一張約 0.5 MB）
const DRIVE_API_DAILY = 10000
/** 今天用了幾次的鍵（Google 的每日次數照美國太平洋時間重算） */
function driveApiDayKey() {
  return 'DRIVE_API_N_' + Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyyMMdd')
}
function driveApiOn() {
  try {
    const c = CacheService.getScriptCache()
    return !c.get('NO_DRIVE_API') && Number(c.get(driveApiDayKey()) || 0) < DRIVE_API_DAILY
  } catch (e) {
    return true
  }
}
/** Google 那邊有沒有開 Drive API（沒試過、或關掉的原因不是「沒開」都算有） */
function driveApiEnabled() {
  try {
    return CacheService.getScriptCache().get('NO_DRIVE_API') !== 'off'
  } catch (e) {
    return true
  }
}
/** 關掉同時讀寫一陣子：res＝Drive API 的回覆（只有「沒開、沒權限」才關）；沒有 res＝整批送不出去 */
function driveApiOff(res, seconds) {
  // 一時忙線（429、500、403 太頻繁）不關，下次照樣試
  const text = res ? String(res.getContentText()).slice(0, 2000) : ''
  if (res && !(res.getResponseCode() === 403 && /accessNotConfigured|SERVICE_DISABLED|has not been used|is disabled|insufficient authentication scopes|ACCESS_TOKEN_SCOPE_INSUFFICIENT/i.test(text))) return
  try {
    // off＝Google 那邊沒開（App 會教擁有者打開）；busy＝次數用完或送不出去，等一下自己恢復
    CacheService.getScriptCache().put('NO_DRIVE_API', res ? 'off' : 'busy', seconds || 21600)
  } catch (e) {
    /* 記不住就算了 */
  }
}
/** 同時送一批請求；Drive API 用不了就回 null（呼叫的人改用一個一個來） */
function fetchMany(requests) {
  if (!requests.length || !driveApiOn()) return null
  try {
    const c = CacheService.getScriptCache()
    const k = driveApiDayKey()
    c.put(k, String(Number(c.get(k) || 0) + requests.length), 21600)
  } catch (e) {
    /* 算不了就算了 */
  }
  const token = ScriptApp.getOAuthToken()
  requests.forEach(function (r) {
    r.headers = Object.assign({ Authorization: 'Bearer ' + token }, r.headers || {})
    r.muteHttpExceptions = true
  })
  try {
    return UrlFetchApp.fetchAll(requests)
  } catch (e) {
    // 每日次數用完：關 6 小時；其他（例如其中一個逾時）：關 10 分鐘
    driveApiOff(null, /too many times|urlfetch|bandwidth|quota/i.test(String((e && e.message) || e)) ? 21600 : 600)
    return null
  }
}
/** 讀很多個檔案：回傳 { 檔案ID: 內容文字 }，讀不到的是 null */
function readFiles(ids) {
  const out = {}
  for (let at = 0; at < ids.length; at += FETCH_CHUNK) {
    const part = ids.slice(at, at + FETCH_CHUNK)
    const res = fetchMany(
      part.map(function (id) {
        return { url: DRIVE_API + '/' + encodeURIComponent(id) + '?alt=media', method: 'get' }
      }),
    )
    part.forEach(function (id, i) {
      const r = res && res[i]
      if (r && r.getResponseCode() === 200) {
        out[id] = r.getContentText('UTF-8')
        return
      }
      if (r) driveApiOff(r)
      try {
        out[id] = DriveApp.getFileById(id).getBlob().getDataAsString()
      } catch (e) {
        out[id] = null
      }
    })
  }
  return out
}
/** Drive API 建一個檔（multipart：先放檔名、資料夾，再放內容） */
function uploadRequest(P, it) {
  const boundary = 'b' + Utilities.getUuid().replace(/-/g, '')
  const type = 'multipart/related; boundary=' + boundary
  const meta = JSON.stringify({ name: it.name, mimeType: 'application/json', parents: [P.SYNC_FOLDER] })
  const body = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + it.text + '\r\n--' + boundary + '--'
  // 直接給 Blob（不要轉成數字陣列：照片一張就是幾十萬個元素，很吃記憶體）
  return { url: DRIVE_UPLOAD, method: 'post', contentType: type, payload: Utilities.newBlob(body, type) }
}
/**
 * 建很多個檔案：items＝[{ name, text }]，回傳一樣順序的檔案 ID。
 * 有一個建不起來就丟錯誤（整批不算，App 下次重傳）；這次已經建好的先丟到垃圾桶，不留沒人用的檔案。
 */
function createFiles(P, items) {
  const ids = []
  try {
    let at = 0
    while (at < items.length) {
      // 一組最多 20 個、約 4 MB
      let end = at
      let bytes = 0
      while (end < items.length && end - at < FETCH_CHUNK && (end === at || bytes + items[end].text.length <= UPLOAD_GROUP_BYTES)) bytes += items[end++].text.length
      const part = items.slice(at, end)
      at = end
      const res = P.SYNC_FOLDER && driveApiOn() ? fetchMany(part.map((it) => uploadRequest(P, it))) : null
      part.forEach(function (it, i) {
        const r = res && res[i]
        let id = ''
        if (r && r.getResponseCode() === 200) {
          try {
            id = JSON.parse(r.getContentText()).id || ''
          } catch (e) {
            id = ''
          }
        } else if (r) driveApiOff(r)
        if (!id) id = syncFolder(P).createFile(Utilities.newBlob(it.text, 'application/json', it.name)).getId()
        ids.push(id)
      })
    }
  } catch (e) {
    trashFiles(ids)
    throw e
  }
  return ids
}
/** 把很多個舊檔案丟到垃圾桶（丟不掉就算了，不影響資料） */
function trashFiles(ids) {
  for (let at = 0; at < ids.length; at += FETCH_CHUNK) {
    const part = ids.slice(at, at + FETCH_CHUNK)
    const res = fetchMany(
      part.map(function (id) {
        return { url: DRIVE_API + '/' + encodeURIComponent(id), method: 'patch', contentType: 'application/json', payload: '{"trashed":true}' }
      }),
    )
    part.forEach(function (id, i) {
      const r = res && res[i]
      if (r && (r.getResponseCode() === 200 || r.getResponseCode() === 404)) return
      if (r) driveApiOff(r)
      try {
        DriveApp.getFileById(id).setTrashed(true)
      } catch (e) {
        /* 檔案已經不在了 */
      }
    })
  }
}

/** 編輯者一次最多刪幾次盤點（「刪除全部盤點紀錄」只有擁有者、管理員能做） */
const EDITOR_MAX_SESSION_DELETES = 20
/**
 * 上傳：比較新的才寫（以裝置上的修改時間為準）；舊檔案丟到垃圾桶。
 * manager＝擁有者或管理員。編輯者送來的這些會被忽略（放在 ignored 回給 App）：
 * - 正航產品表（erp:…）：匯入會改大家的帳面數
 * - 一次刪超過 20 次盤點（等於刪光大家的紀錄）
 * - 刪除點貨紀錄（pick:…）：編輯者可以新增、修改，不能刪（4.7.1）
 * 每筆的修改時間 t 最多只能比現在晚 10 分鐘。
 * 共用設定（settings）裡的盲盤、複盤規則：編輯者送來的不算，保留雲端原本的（儲位、品項清單照常收）。
 */
function syncPush(data, props, P, manager) {
  const lock = LockService.getScriptLock()
  lock.waitLock(30000)
  try {
    // 鎖住後重新讀資料夾、目錄檔的位置（兩台同時第一次上傳，才不會各建一個資料夾）
    ;['SYNC_FOLDER', 'SYNC_INDEX'].forEach(function (k) {
      const v = props.getProperty(k)
      if (v) P[k] = v
      else delete P[k]
    })
    const idx = readIndex(P)
    if (!P.SYNC_FOLDER) syncFolder(P) // 第一次同步：先建資料夾（新檔案要放在裡面）
    // 順序號一定越來越大（同一毫秒兩次上傳也不會重複），下載時才不會漏；鎖住後重新讀一次，別人剛寫的才不會被蓋掉
    let seq = Math.max(Date.now(), Number(props.getProperty('SYNC_SEQ') || 0) + 1)
    let n = 0
    const skipped = []
    const ignored = []
    const isSessionDel = function (r) {
      return !!r && r.del && String(r.k).indexOf('session:') === 0
    }
    const tooManyDeletes =
      !manager &&
      (data.records || []).filter(function (r) {
        return isSessionDel(r)
      }).length > EDITOR_MAX_SESSION_DELETES
    // 4.7.6：先決定每一筆要怎麼做，最後再一次「同時」建新檔、丟舊檔（以前一筆一筆來，幾十筆要等快一分鐘）。
    // 順序也改了：新檔都建好、目錄寫好，才把舊檔丟到垃圾桶（以前先丟再建，建失敗會少一筆）
    const toTrash = []
    const toWrite = [] // { entry: 目錄裡那一格, name, text, skip }
    const writing = {} // 同一批裡同一個鍵出現兩次：前面那次就不用建檔了
    const trash = function (e) {
      if (e && e.f) toTrash.push(e.f)
    }
    const unwrite = function (k) {
      if (writing[k]) writing[k].skip = true
      delete writing[k]
    }
    // 修改時間是裝置自己填的：最多只能比現在晚 10 分鐘（手機時鐘差一點沒關係），
    // 不然填一個很遠的未來時間，之後連管理員都蓋不掉、刪不掉（4.7.1 code review）
    const maxT = Date.now() + 10 * 60 * 1000
    ;(data.records || []).forEach(function (r) {
      if (!r || typeof r.k !== 'string' || !r.k) return
      if (!r.del && r.d === undefined) return // 沒有內容、也不是刪除：略過，不讓整批失敗
      r.t = Number(r.t)
      if (!(r.t > 0)) r.t = Date.now()
      if (r.t > maxT) r.t = maxT
      const cur = idx.map[r.k]
      // 照片不會改：雲端已經有這張就不用再存一次（算成功，App 才會記成「傳過了」）
      if (isPhotoKey(r.k) && !r.del && cur && cur.f && !cur.del) {
        n++
        return
      }
      // 只有擁有者、管理員能改的：編輯者送來的不寫（點貨紀錄：編輯者可以新增、修改，不能刪）
      if (!manager && (String(r.k).indexOf('erp:') === 0 || (tooManyDeletes && isSessionDel(r)) || (r.del && r.k.indexOf('pick:') === 0))) {
        ignored.push(r.k)
        return
      }
      if (cur && cur.t > r.t) {
        skipped.push(r.k) // 雲端已經有比較新的（別台改的）
        return
      }
      // 同一台重送同一個版本（上傳逾時、App 沒收到回覆又送一次）：已經收過了，不用再建一次檔
      if (!r.del && cur && cur.f && !cur.del && cur.t === r.t && cur.d === String(data.dev || '')) {
        n++
        return
      }
      // 編輯者改了儲位、品項清單：盲盤、複盤規則照雲端原本的（沒有就不帶，別台就不會跟著改）
      if (r.k === 'settings' && !manager && !r.del && r.d && typeof r.d === 'object') {
        let old = {}
        if (cur && cur.f && !cur.del) {
          try {
            old = JSON.parse(DriveApp.getFileById(cur.f).getBlob().getDataAsString()) || {}
          } catch (err) {
            old = {}
          }
        }
        r.d = Object.assign({}, r.d)
        ;['blind', 'recount'].forEach(function (f) {
          if (old[f] != null) r.d[f] = old[f]
          else delete r.d[f]
        })
      }
      trash(cur)
      unwrite(r.k)
      const entry = { t: r.t, s: seq++, d: String(data.dev || ''), f: '', del: !!r.del }
      if (!r.del) {
        const text = JSON.stringify(r.d)
        if (isPhotoKey(r.k)) entry.z = text.length // 照片記大小：下載時先算好一次拿幾張，不用讀了又丟
        toWrite.push((writing[r.k] = { entry: entry, name: r.k.replace(/[^\w-]/g, '_') + '.json', text: text }))
      }
      idx.map[r.k] = entry
      // 刪掉一次盤點：它的照片也一起丟掉
      if (r.del && String(r.k).indexOf('session:') === 0) {
        const prefix = 'photo:' + String(r.k).slice('session:'.length) + ':'
        Object.keys(idx.map).forEach(function (k) {
          if (k.indexOf(prefix) !== 0 || idx.map[k].del) return
          trash(idx.map[k])
          unwrite(k)
          idx.map[k] = { t: r.t, s: seq++, d: String(data.dev || ''), f: '', del: true }
        })
      }
      n++
    })
    const jobs = toWrite.filter(function (w) {
      return !w.skip
    })
    const ids = createFiles(P, jobs)
    jobs.forEach(function (w, i) {
      w.entry.f = ids[i]
    })
    writeIndex(P, idx)
    props.setProperty('SYNC_SEQ', String(seq))
    trashFiles(toTrash)
    const why = []
    if (tooManyDeletes) why.push('一次刪超過 ' + EDITOR_MAX_SESSION_DELETES + ' 次盤點只有擁有者或管理員能做：這些刪除沒有同步，雲端和別台的紀錄還在')
    if (ignored.some(function (k) {
      return String(k).indexOf('erp:') === 0
    }))
      why.push('正航產品表只有擁有者或管理員能匯入：這次沒有同步給大家')
    if (ignored.some(function (k) {
      return String(k).indexOf('pick:') === 0
    }))
      why.push('點貨紀錄只有擁有者或管理員能刪：這些刪除沒有同步，雲端和別台的紀錄還在')
    return { ok: true, n: n, skipped: skipped, ignored: ignored, ignoredMsg: why.join('；') }
  } finally {
    lock.releaseLock()
  }
}

/**
 * 下載：別台送來、比 since 新的；一次最多約 8 MB、150 筆（同時讀用不了時 60 筆）或 20 秒，more＝還有。
 * 4.7.6：檔案同時讀，一次可以給比較多筆；限時 20 秒，手機切到背景、網路斷掉時要重來的比較少
 */
const PULL_BUDGET_MS = 20000
function syncPull(data, P) {
  const started = Date.now()
  const PULL_MAX = driveApiOn() ? 150 : 60
  const idx = readIndex(P)
  const since = Number(data.since) || 0
  const dev = String(data.dev || '')
  const all = Object.keys(idx.map)
    .map(function (k) {
      return { k: k, e: idx.map[k] }
    })
    .filter(function (x) {
      return x.e.s > since && !isPhotoKey(x.k) // 照片不在這裡給（App 打開那次盤點才用 photos 要）
    })
    .sort(function (a, b) {
      return a.e.s - b.e.s
    })
  const records = []
  let size = 0
  let next = since
  let i = 0 // all 裡第一筆還沒給的
  const full = function () {
    return records.length && (size > 8e6 || records.length >= PULL_MAX || Date.now() - started > PULL_BUDGET_MS)
  }
  while (i < all.length && !full()) {
    // 這一輪：往後拿到 FETCH_CHUNK 筆要給的，一起讀（自己送的不用再下載，不算）
    const group = []
    let want = 0
    for (let j = i; j < all.length && want < FETCH_CHUNK && records.length + want < PULL_MAX; j++) {
      group.push(all[j])
      if (all[j].e.d !== dev) want++
    }
    const texts = readFiles(
      group
        .filter(function (x) {
          return x.e.d !== dev && !x.e.del && x.e.f
        })
        .map(function (x) {
          return x.e.f
        }),
    )
    for (const x of group) {
      if (full()) break
      next = x.e.s
      i++
      if (x.e.d === dev) continue
      let d = null
      const text = !x.e.del && x.e.f ? texts[x.e.f] : null
      if (text != null) {
        size += text.length
        try {
          d = JSON.parse(text)
        } catch (e) {
          d = null
        }
      }
      records.push({ k: x.k, t: x.e.t, del: !!x.e.del, d: d })
    }
  }
  return { ok: true, records: records, next: next, more: i < all.length }
}

/**
 * 要照片：給一串 photo:盤點ID:照片ID，回傳有的那些（一次最多約 8 MB 或 20 秒，more＝還有沒給的）。
 * 4.7.6：照目錄記的大小（舊照片沒記，當 0.6 MB）先算好這一輪拿幾張（最多 12 張），一起讀
 */
function syncPhotos(data, P) {
  const started = Date.now()
  const idx = readIndex(P)
  const keys = (data.keys || []).filter(isPhotoKey).slice(0, 200)
  const records = []
  let size = 0
  let i = 0
  const full = function () {
    return records.length && (size > 8e6 || Date.now() - started > PULL_BUDGET_MS)
  }
  while (i < keys.length && !full()) {
    const group = []
    let plan = size
    for (let j = i; j < keys.length && group.length < 12 && (!group.length || plan <= 8e6); j++) {
      group.push(keys[j])
      const e = idx.map[keys[j]]
      if (e && !e.del && e.f) plan += e.z || 6e5
    }
    const texts = readFiles(
      group
        .map(function (k) {
          return idx.map[k]
        })
        .filter(function (e) {
          return e && !e.del && e.f
        })
        .map(function (e) {
          return e.f
        }),
    )
    for (const k of group) {
      if (full()) break
      i++
      const e = idx.map[k]
      if (!e || e.del || !e.f) {
        records.push({ k: k, d: null }) // 雲端沒有這張（App 不用再要）
        continue
      }
      let d = null
      const text = texts[e.f]
      if (text != null) {
        size += text.length
        try {
          d = JSON.parse(text)
        } catch (err) {
          d = null
        }
      }
      records.push({ k: k, d: d })
    }
  }
  return { ok: true, records: records, more: i < keys.length }
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
  const p = sh.protect().setDescription('由聖佳智慧庫存 App 自動產生：請在 App 裡改').setWarningOnly(true)
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
  const BASE = ['料號', '品名', '品牌', '型號', '尺寸／規格', '實盤', '帳面', '差異', '差異比例', '在哪裡（位置 數量）', '最近盤點', '盤點人', '複盤']
  const NOTE = ['原因說明（請填）', '處理方式（請填）', '主管確認']
  const sh = sheetAt(ss, '盤差報告', 3)
  // 先把主管填過的留下來（用料號對；用標題找欄位，以後加欄位也不會錯位）
  const notes = {}
  const last = sh.getLastRow()
  const lastCol = sh.getLastColumn()
  if (last > 1 && lastCol > 0) {
    const old = sh.getRange(1, 1, last, lastCol).getValues()
    const at = NOTE.map(function (h) {
      return old[0].indexOf(h)
    })
    old.slice(1).forEach(function (r) {
      if (r[0])
        notes[String(r[0])] = at.map(function (i, k) {
          return i >= 0 ? r[i] : k === 2 ? false : ''
        })
    })
  }
  const pick = function (r, h) {
    return col(h) >= 0 ? r[col(h)] : ''
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
      // 原因說明：主管填過的優先；沒填就帶 App 複盤時選的原因
      const note = (notes[String(r[col('料號')])] || ['', '', false]).slice()
      if (!note[0] && pick(r, '差異原因')) note[0] = pick(r, '差異原因')
      return [r[col('料號')], r[col('品名')], r[col('品牌')], r[col('型號')], r[col('尺寸／規格')], r[col('實盤')], r[col('帳面')], d, book ? d / book : '', r[col('在哪裡（位置 數量）')], cellOf(r[col('最近盤點')]), r[col('盤點人')], pick(r, '複盤')].concat(note)
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
  // 後三欄給主管填，不跳警告（欄位字母照欄數算）
  const letter = function (n) {
    return String.fromCharCode(64 + n)
  }
  protectWarn(sh, letter(BASE.length + 1) + ':' + letter(width))
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
  sh.getRange('E2').setValue('這一頁由聖佳智慧庫存 App 自動產生，不用改；細節看後面的工作表。').setFontColor(COLOR.gray)
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
      // 現在剩幾個：實盤、帳面都有就取比較少的（跟 App 的叫貨提醒一樣）
      const book = r[col('帳面')]
      const real = r[col('實盤')]
      const counted = col('最近盤點') >= 0 && r[col('最近盤點')] !== '' // 沒盤過的，實盤是 0 但不能算
      const now = book !== '' && counted ? Math.min(Number(book), Number(real)) : book !== '' ? book : real
      return [r[col('料號')], r[col('品名')], r[col('尺寸／規格')], now, r[col('剩幾個要叫貨（安全庫存）')]]
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
