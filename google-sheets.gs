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
 * - 擁有者的手機不見了：Apps Script 左邊「專案設定」→ 最下面「指令碼屬性」→ 把 SYNC_KEY 的值複製起來（不要刪掉它），貼到 App 新手機的「貼上你備份的連結」欄，就接回來了。
 *
 * 共用 AI 金鑰：擁有者在 App「設定 → AI → 放到雲端」→ 金鑰存在這裡的指令碼屬性 CLAUDE_KEY。
 * 同事拍照時，請求送到這裡，由這裡拿金鑰去問 Claude，同事的手機不會有金鑰；移除權限的人就不能再用。
 * 費用上限（Claude 沒有免費額度，擁有者付錢）：只准 Haiku／Sonnet（Opus 要 CLAUDE_ALLOW_OPUS＝1）、每人每天 100 次（CLAUDE_DAILY_LIMIT）、一次最多 8 MB、不准工具。
 * 編輯者不能改盲盤／複盤規則、不能匯入正航產品表、不能一次刪光盤點紀錄（雲端也擋）。
 * 更新方法：整段重新貼上 → 存檔 →「部署」→「管理部署作業」→ 鉛筆 → 版本選「新版本」→ 部署（網址不會變）。
 *
 * 4.11：用 Google 登入（OpenID Connect）：擁有者在 App 設定貼上「Google 登入用的 Client ID」（指令碼屬性 GOOGLE_CLIENT_ID，不是密碼）。
 * - 同事按「用 Google 登入」→ 驗證過的信箱在「預先登記」名單（GOOGLE_ALLOW）就直接加入；不在就排進待核准（核准畫面看得到信箱）。
 * - 登入資料綁在「這支手機自己產生的連結碼」：Google 的 nonce ＝ SHA-256(連結碼)，偷到登入資料也不能拿去註冊別支手機。
 * - 擁有者綁定自己的 Google（OWNER_EMAIL）後，換手機用 Google 登入就能接回來（那支手機的連結碼記在 OWNER_HS）。
 * - 一個人可以有好幾支手機（成員的 hs，最多 5 支）；被移除時全部一起失效。
 * 4.10：店裡通行碼＋待核准（同事沒有 LINE、用 Android、擁有者請假時也能加入）：
 * - 擁有者／管理員在 App 設一組通行碼（指令碼屬性 JOIN_CODE）；同事在 App 填名字＋通行碼 → 進「待核准」（JOIN_PENDING，最多 20 人、7 天過期）；
 *   管理員在 App 按「核准」才會變成共用名單裡的人（預設檢視者）。猜錯 10 次，全部暫停 10 分鐘（JOIN_FAILS）。
 * - 申請的人自己在手機產生連結碼，雲端只存雜湊值；密碼不用傳給任何人。
 * 4.9：拿掉 Gemini。舊的共用 Gemini 金鑰（指令碼屬性 GEMINI_KEY）和它的用量紀錄、次數上限（AI_USAGE、AI_DAILY_LIMIT），部署新版本後第一次有人連線就自動刪掉。
 *
 * 4.7.6：同步變快——很多筆資料、照片「同時」讀寫（Drive API），不再一筆一筆來；
 * 記住目錄檔的位置，不用每次搜尋。Drive API 用不了會自動改回舊方法（結果一樣，只是比較慢），
 * App 的設定頁會提醒擁有者：左邊「服務」→ 加入「Drive API」→ 再部署一次新版本，就會打開。
 * 第一次更新到這一版，Google 可能會再問一次授權，按允許即可。
 */
const RAW = '盤點紀錄'
const HEAD = ['盤點日期', '時間', '位置', '品名', '品牌', '型號', '尺寸／規格', '數量', '照片框數', '來源', '盤點ID', '盤點人']
const SYNC_ACTIONS = ['push', 'pull', 'photos', 'members', 'invite', 'remove', 'setRole', 'rename', 'reissue', 'joinInfo', 'joinCodeSet', 'joinApprove', 'joinReject', 'googleAllow', 'googleUnallow', 'googleSetup']
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
  // 4.9 拿掉 Gemini：不用的金鑰留著有外流風險，自動刪掉（使用者 10/19 同意）
  for (const k of ['GEMINI_KEY', 'AI_USAGE', 'AI_DAILY_LIMIT'])
    if (P[k]) {
      props.deleteProperty(k)
      delete P[k]
    }
  // 還沒開啟同步：只有使用者在 App 按「開啟同步」（hello）才能設定擁有者；背景同步不能設定，也讀不到資料
  // 設定擁有者（第一台裝置）：一定要 init:true（只有 App 的「開啟同步」會帶）。曾經設過擁有者（SYNC_EVER）之後，要有人在 Apps Script
  // 親手加指令碼屬性 ALLOW_REINIT＝1 才能重來。以前「沒有 SYNC_KEY 時，任何人的 hello 都會變成擁有者」：網址放進公開的 App 之後太危險，
  // 而且等核准的手機每 40 秒就會自己打一次 hello。
  if (P.SYNC_KEY && !P.SYNC_EVER) {
    props.setProperty('SYNC_EVER', '1')
    P.SYNC_EVER = '1'
  }
  if (!P.SYNC_KEY && data.action === 'hello' && data.init === true && /^[\w-]{16,64}$/.test(String(data.key || '')) && (!P.SYNC_EVER || P.ALLOW_REINIT === '1')) {
    props.setProperty('SYNC_KEY', String(data.key))
    props.setProperty('SYNC_EVER', '1')
    props.deleteProperty('ALLOW_REINIT')
    P.SYNC_KEY = String(data.key)
    P.SYNC_EVER = '1'
    delete P.ALLOW_REINIT
  }
  const syncOn = !!P.SYNC_KEY
  // 同事用店裡通行碼申請加入：不用連結碼（他還沒有），所以放在檢查連結碼之前
  if (data.action === 'joinRequest') return json(syncOn ? joinRequest(data, props, P) : { ok: false, error: '還沒開啟同步：請先由擁有者在第一台裝置開啟' })
  // 用 Google 登入（不用連結碼）；googleConfig 只回「Client ID」（公開資訊，App 靠它決定要不要顯示「用 Google 登入」）
  if (data.action === 'googleConfig') return json({ ok: true, clientId: P.GOOGLE_CLIENT_ID || '' })
  if (data.action === 'joinGoogle') return json(syncOn ? joinGoogle(data, props) : { ok: false, error: '還沒開啟同步：請先由擁有者在第一台裝置開啟' })
  const who = syncOn ? whoIs(P, data.key) : null
  if (syncOn && !who) {
    // 申請了、還沒被核准：告訴 App「等一下」，不要當成連結碼錯誤；取消申請（joinWithdraw）就把那一筆拿掉
    const hk = sha(data.key || '')
    if (
      pendingList(P.JOIN_PENDING).some(function (x) {
        return x.h === hk
      })
    ) {
      if (data.action === 'joinWithdraw') {
        const lock = LockService.getScriptLock()
        lock.waitLock(20000)
        try {
          props.setProperty(
            'JOIN_PENDING',
            JSON.stringify(
              pendingList(props.getProperty('JOIN_PENDING')).filter(function (x) {
                return x.h !== hk
              }),
            ),
          )
        } finally {
          lock.releaseLock()
        }
        return json({ ok: true })
      }
      return json({ ok: false, pending: true, error: '已送出申請，等管理員核准後就能用' })
    }
    const revoked = JSON.parse(P.SYNC_REVOKED || '[]').indexOf(sha(data.key || '')) >= 0
    return json({ ok: false, revoked: revoked, error: revoked ? '這台已經被移除權限，不能再同步' : '連結碼不對或已經失效：請找管理員要一個新的連結碼' })
  }
  if (!syncOn && P.SYNC_EVER) return json({ ok: false, error: '雲端的擁有者連結碼（指令碼屬性 SYNC_KEY）不見了：請擁有者把它補回去；真的要重來，在指令碼屬性加 ALLOW_REINIT＝1，再從 App 按「第一次設定」' })
  if (!syncOn && (data.action === 'hello' || SYNC_ACTIONS.indexOf(data.action) >= 0)) return json({ ok: false, error: '還沒開啟同步：請在第一台裝置按「開啟多人同步」' })
  if (who) touchSeen(props, P, who)
  if (who && data.action === 'bindGoogle') return json(bindGoogle(data, props, who))
  // 搬家（例如從個人帳號搬到公司帳號）：只有擁有者能做；搬完後，舊的位置會告訴每台 App 新網址，大家自動跟過去
  if (who && who.role === 'owner' && data.action === 'exportMembers') return json({ ok: true, members: membersOf(P), revokedHashes: JSON.parse(P.SYNC_REVOKED || '[]'), ownerName: P.SYNC_OWNER_NAME || '', extra: { JOIN_CODE: P.JOIN_CODE || '', GOOGLE_CLIENT_ID: P.GOOGLE_CLIENT_ID || '', GOOGLE_ALLOW: P.GOOGLE_ALLOW || '', OWNER_EMAIL: P.OWNER_EMAIL || '', OWNER_HS: P.OWNER_HS || '' } })
  if (who && who.role === 'owner' && data.action === 'importMembers') {
    props.setProperty('SYNC_MEMBERS', JSON.stringify(data.members || []))
    props.setProperty('SYNC_REVOKED', JSON.stringify(data.revokedHashes || []))
    if (data.ownerName) props.setProperty('SYNC_OWNER_NAME', String(data.ownerName))
    // 通行碼、Google 登入的設定也一起搬（只收這幾個、而且都是字串）
    ;['JOIN_CODE', 'GOOGLE_CLIENT_ID', 'GOOGLE_ALLOW', 'OWNER_EMAIL', 'OWNER_HS'].forEach(function (k) {
      const v = data.extra && data.extra[k]
      if (typeof v === 'string' && v && v.length < 8000) props.setProperty(k, v)
    })
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
  const pend = canManage ? pendingList(P.JOIN_PENDING).length : 0
  // 名單（只有名字，給每台 App 顯示「誰盤的」、選盤點人；名字只有擁有者／管理員能改）
  const roster = who
    ? [{ id: 'owner', name: P.SYNC_OWNER_NAME || '擁有者' }].concat(
        membersOf(P).map(function (m) {
          return { id: m.id, name: m.name }
        }),
      )
    : []
  // 共用 AI 金鑰：aiClaude＝擁有者有沒有放 Claude 的共用金鑰（指令碼屬性 CLAUDE_KEY）；每個回覆都帶，App 才知道能不能用
  if (!who && (data.action === 'ai' || data.action === 'aiKey')) return json({ ok: false, error: '還沒開啟同步，不能用共用的 AI 金鑰' })
  const aiClaude = !!P.CLAUDE_KEY
  // 4.9 只用 Claude：舊版 App 要放、要用 Gemini 金鑰的都擋掉（不然 Gemini 金鑰會被存成 CLAUDE_KEY）
  const oldApp = '這個 App 是舊版（還在用 Gemini）：請在 App「設定」最下面按「檢查更新」，更新後改用 Claude'
  if (who && data.action === 'aiKey') {
    if (who.role !== 'owner') return json({ ok: false, aiClaude: aiClaude, pend: pend, me: me, error: '只有擁有者可以設定共用的 AI 金鑰' })
    if (data.provider !== 'claude') return json({ ok: false, aiClaude: aiClaude, pend: pend, me: me, error: oldApp })
    if (data.aiKey) props.setProperty('CLAUDE_KEY', String(data.aiKey))
    else props.deleteProperty('CLAUDE_KEY')
    return json({ ok: true, aiClaude: !!data.aiKey, me: me })
  }
  if (who && data.action === 'ai') {
    if (!canEdit) return json({ ok: false, viewer: true, aiClaude: aiClaude, pend: pend, me: me, error: '你是檢視者，只能看' })
    if (data.provider !== 'claude') return json({ ok: false, status: 403, aiClaude: aiClaude, pend: pend, me: me, error: oldApp })
    return json(Object.assign(claudeProxy(data, P, props, who), { aiClaude: aiClaude, pend: pend, me: me }))
  }
  // ver 2：盤點的照片分開存（photo:盤點ID:照片ID），下載只拿資料，照片用 photos 另外要；App 看到 ver 才改用新格式上傳
  // fast：有沒有用到「同時讀寫」（4.7.6）；沒有的話，App 會教擁有者打開 Drive API
  if (data.action === 'hello') return json({ ok: true, ver: SYNC_VER, aiClaude: aiClaude, pend: pend, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), me: me, roster: roster, fast: driveApiEnabled() })
  if (data.action === 'pull') return json(Object.assign(syncPull(data, P), { ver: SYNC_VER, aiClaude: aiClaude, pend: pend, me: me, roster: roster, fast: driveApiEnabled() }))
  if (data.action === 'photos') return json(Object.assign(syncPhotos(data, P), { ver: SYNC_VER, aiClaude: aiClaude, pend: pend, me: me }))
  if (data.action === 'push') return json(canEdit ? Object.assign(syncPush(data, props, P, !!canManage), { ver: SYNC_VER, aiClaude: aiClaude, pend: pend, me: me }) : { ok: false, viewer: true, aiClaude: aiClaude, pend: pend, me: me, error: '你是檢視者，只能看' })
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

/** 一次最多 8 MB（一張照片＋30 張樣品照大約 2～4 MB） */
const AI_MAX_BYTES = 8 * 1024 * 1024
/**
 * 共用 Claude 金鑰（4.8）：幫 App 問 Claude（金鑰只在這裡，指令碼屬性 CLAUDE_KEY）。只准兩種請求：列出模型、看圖回答（v1/messages）。
 * Claude 沒有免費額度（擁有者付錢），所以有上限：
 * - 每人每天最多 100 次（指令碼屬性 CLAUDE_DAILY_LIMIT 可以改）、一次最多 8 MB
 * - 模型只准 Haiku、Sonnet；Opus 比較貴，要在指令碼屬性加 CLAUDE_ALLOW_OPUS＝1 才放行
 * - 一次最多回 16000 個字詞；只准一則訊息、內容只能是文字和照片；不准工具（網路搜尋、程式執行等）
 */
const CLAUDE_MAX_OUTPUT = 16000
const CLAUDE_DAILY_LIMIT = 100
function claudeLimit(P) {
  return Number(P.CLAUDE_DAILY_LIMIT) > 0 ? Number(P.CLAUDE_DAILY_LIMIT) : CLAUDE_DAILY_LIMIT
}
function claudeModelOk(name, P) {
  const m = String(name || '')
  if (/^claude-opus-/.test(m) && !(P && P.CLAUDE_ALLOW_OPUS === '1')) return false
  return /^claude-(haiku|sonnet|opus)-[a-z0-9.\-]+$/.test(m)
}
function claudeProxy(data, P, props, who) {
  if (!P.CLAUDE_KEY) return { ok: false, status: 403, error: '擁有者還沒把 Claude 金鑰放到雲端（或已經停止共用）' }
  const path = typeof data.path === 'string' ? data.path : ''
  const list = /^v1\/models(\?limit=\d+)?$/.test(path)
  const msg = path === 'v1/messages'
  if (!list && !msg) return { ok: false, status: 400, error: '不支援的 AI 請求' }
  const opt = { method: 'get', headers: { 'x-api-key': P.CLAUDE_KEY, 'anthropic-version': '2023-06-01' }, muteHttpExceptions: true }
  if (msg) {
    const raw = typeof data.body === 'string' ? data.body : JSON.stringify(data.body || {})
    if (raw.length > AI_MAX_BYTES) return { ok: false, status: 400, error: '這次送的照片太大（超過 8 MB）：請少放一些樣品照再試' }
    let src
    try {
      src = JSON.parse(raw)
    } catch (e) {
      src = null
    }
    if (!src || typeof src !== 'object' || !Array.isArray(src.messages)) return { ok: false, status: 400, error: 'AI 請求的格式不對' }
    if (!claudeModelOk(src.model, P)) return { ok: false, status: 404, error: '共用金鑰不能用這個模型（' + src.model + '）：只能用 Claude 的 Haiku、Sonnet' + (P.CLAUDE_ALLOW_OPUS === '1' ? '、Opus' : '（Opus 比較貴，擁有者沒有開放）') }
    // 只准一則使用者訊息，內容只能是文字和直接附上的照片（不准用網址叫 Claude 去抓東西、不准附文件）
    const m0 = src.messages.length === 1 && src.messages[0] && src.messages[0].role === 'user' && Array.isArray(src.messages[0].content) ? src.messages[0] : null
    const okBlocks =
      m0 &&
      m0.content.length > 0 &&
      m0.content.every(function (b) {
        return b && ((b.type === 'text' && typeof b.text === 'string') || (b.type === 'image' && b.source && b.source.type === 'base64'))
      })
    if (!okBlocks) return { ok: false, status: 400, error: 'AI 請求的內容不對（只能是文字和照片）' }
    // 只留需要的欄位：結構化輸出（JSON 格式）、思考設定；其他（工具、網路搜尋、程式執行、MCP…）一律拿掉
    const body = { model: src.model, max_tokens: Math.min(Number(src.max_tokens) || CLAUDE_MAX_OUTPUT, CLAUDE_MAX_OUTPUT), messages: [{ role: 'user', content: m0.content }] }
    const oc = src.output_config && typeof src.output_config === 'object' ? src.output_config : {}
    const outCfg = {}
    if (oc.format && oc.format.type === 'json_schema' && oc.format.schema) outCfg.format = { type: 'json_schema', schema: oc.format.schema }
    if (oc.effort === 'low' || oc.effort === 'medium') outCfg.effort = oc.effort
    if (outCfg.format || outCfg.effort) body.output_config = outCfg
    const th = src.thinking && src.thinking.type
    if (th === 'between_tools' || th === 'disabled') body.thinking = { type: th }
    const used = aiCount(props, P, who, 'AI_USAGE_CLAUDE', claudeLimit(P))
    if (used < 0) return { ok: false, status: 429, quota: true, error: '今天用共用 Claude 的次數已經到上限（每人每天 ' + claudeLimit(P) + ' 次），明天會自動恢復；急用請找擁有者' }
    opt.method = 'post'
    opt.contentType = 'application/json'
    opt.payload = JSON.stringify(body)
  }
  let res
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/' + path, opt)
  } catch (e) {
    // 連不到 Claude、或等太久（Apps Script 最多等大約 1 分鐘）：回 408，App 會改用比較快的 Haiku，
    // 不要回「太多人用」讓 App 馬上用同一個模型再送一次（Claude 那邊可能還在做、照樣計費）
    return { ok: true, status: 408, body: JSON.stringify({ type: 'error', error: { type: 'timeout_error', message: 'Apps Script 等不到 Claude 回覆：' + String((e && e.message) || e).slice(0, 150) } }) }
  }
  const status = res.getResponseCode()
  let text = res.getContentText()
  // 列模型：只給共用金鑰能用的那些
  if (list && status === 200) {
    try {
      const all = JSON.parse(text)
      all.data = (all.data || []).filter(function (m) {
        return claudeModelOk(m.id, P)
      })
      text = JSON.stringify(all)
    } catch (e) {
      /* 回覆不是 JSON：原樣給 App */
    }
  }
  return { ok: true, status: status, body: text }
}
/** 這個人今天第幾次用共用 AI（記在指令碼屬性 prop）；超過 limit 回 -1（台灣時間每天 0 點重算） */
function aiCount(props, P, who, prop, limit) {
  const today = new Date(Date.now() + 8 * 3600 * 1000).toISOString().slice(0, 10)
  // 用試算表的鎖（跟同步用的鎖分開）：同步上傳照片要鎖比較久，問 AI 不用跟著等
  const lock = (LockService.getDocumentLock && LockService.getDocumentLock()) || LockService.getScriptLock()
  if (!lock.tryLock(10000)) return 0 // 一直等不到（很少見）：這次先不算，不要讓人卡住
  try {
    let usage = {}
    try {
      usage = JSON.parse(props.getProperty(prop) || '{}') || {}
    } catch (e) {
      usage = {}
    }
    if (usage.day !== today) usage = { day: today, n: {} }
    const id = String(who.id)
    const n = (usage.n[id] || 0) + 1
    if (n > limit) return -1
    usage.n[id] = n
    props.setProperty(prop, JSON.stringify(usage))
    return n
  } finally {
    lock.releaseLock()
  }
}

/** App 的「測試連線」會呼叫這個 */
function doGet() {
  return json({ ok: true, app: '拍照盤點' })
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
// ───────────── 店裡通行碼＋待核准 ─────────────
const JOIN_MAX_PENDING = 20 // 等核准的人最多幾個（防止亂灌）
const JOIN_PENDING_DAYS = 7 // 超過幾天沒人核准就過期
const JOIN_MAX_FAILS = 10 // 10 分鐘內猜錯幾次就暫停
const JOIN_FAIL_WINDOW = 10 * 60 * 1000
const JOIN_LOCK_MS = 10 * 60 * 1000
const JOIN_ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789' // 沒有容易看錯的 0／O、1／I／L
/** 通行碼不分大小寫、不管中間有沒有空格或橫線 */
function normCode(x) {
  return String(x == null ? '' : x)
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '')
}
/** 8 個字的新通行碼：用 getUuid 的亂數，超出整除範圍的丟掉（不偏向前面幾個字母） */
function newJoinCode() {
  let out = ''
  while (out.length < 8) {
    const hex = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '')
    for (let i = 0; i + 1 < hex.length && out.length < 8; i += 2) {
      const n = parseInt(hex.substr(i, 2), 16)
      if (n < 248) out += JOIN_ALPHABET.charAt(n % JOIN_ALPHABET.length)
    }
  }
  return out
}
/** 別人填的名字：去掉開頭的 = + - @（寫進試算表會被當成公式），最長 30 字 */
function cleanName(x) {
  return String(x == null ? '' : x)
    .trim()
    .replace(/^[=+\-@\s]+/, '')
    .slice(0, 30)
}
/** 還沒過期的待核准名單 */
function pendingList(raw) {
  const now = Date.now()
  return JSON.parse(raw || '[]').filter(function (x) {
    return now - x.at < JOIN_PENDING_DAYS * 86400000
  })
}
// ───────────── 用 Google 登入 ─────────────
const GOOGLE_ISS = ['https://accounts.google.com', 'accounts.google.com']
const MAX_DEVICES = 5 // 一個人最多幾支手機
/** JWT 的內容（只用來讀 nonce：簽章和有效期限已經由 Google 的 tokeninfo 檢查過） */
function jwtPayload(idToken) {
  return JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(String(idToken).split('.')[1])).getDataAsString())
}
/** 不用打 Google 就能看出不對的登入（格式、給誰的、過期、nonce 對不對）；沒問題回空字串 */
function googleLocalCheck(idToken, newKey, clientId) {
  if (!clientId) return '擁有者還沒開啟 Google 登入'
  const BAD = 'Google 登入的資料不對，請再登入一次'
  const tok = String(idToken || '')
  if (tok.length > 4096 || !/^[\w-]+\.[\w-]+\.[\w-]+$/.test(tok)) return BAD
  let p
  try {
    p = jwtPayload(tok)
  } catch (e) {
    return BAD
  }
  if (p.aud !== clientId) return '這個登入不是給這個 App 的'
  if (!(Number(p.exp) * 1000 > Date.now())) return 'Google 登入已經過期，請再登入一次'
  if (!p.nonce || p.nonce !== sha(newKey)) return '登入和這支手機對不起來，請再登入一次'
  return ''
}
/**
 * 驗證 Google 的 ID token：簽章、有效期限交給 Google 的 tokeninfo；這邊再檢查 aud（是不是給這個 App 的）、iss、信箱有沒有驗證過、
 * nonce 是不是 SHA-256(這支手機自己產生的連結碼)（偷到登入資料也不能拿去註冊別支手機）。成功回 { email, name }，失敗回 { error }
 */
function googleVerify(idToken, newKey, clientId) {
  if (!clientId) return { error: '擁有者還沒開啟 Google 登入' }
  const tok = String(idToken || '')
  if (tok.length > 4096 || !/^[\w-]+\.[\w-]+\.[\w-]+$/.test(tok)) return { error: 'Google 登入的資料不對，請再登入一次' }
  let res
  try {
    res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(tok), { muteHttpExceptions: true })
  } catch (e) {
    return { error: '連不到 Google，請等一下再試' }
  }
  if (res.getResponseCode() !== 200) return { error: 'Google 說這次登入無效或已經過期，請再登入一次' }
  let t
  let nonce
  try {
    t = JSON.parse(res.getContentText())
    nonce = t.nonce || jwtPayload(tok).nonce
  } catch (e) {
    return { error: 'Google 登入的資料不對，請再登入一次' }
  }
  if (GOOGLE_ISS.indexOf(t.iss) < 0 || t.aud !== clientId) return { error: '這個登入不是給這個 App 的' }
  if (!(Number(t.exp) * 1000 > Date.now())) return { error: 'Google 登入已經過期，請再登入一次' }
  if (String(t.email_verified) !== 'true' || !t.email) return { error: '這個 Google 帳號的信箱還沒有驗證' }
  if (!nonce || nonce !== sha(newKey)) return { error: '登入和這支手機對不起來，請再登入一次' }
  const email = String(t.email).toLowerCase()
  // 名字是別人控制的（Google 顯示名稱）：去掉開頭的公式符號，不然盤點人寫進試算表會被當成公式
  return { email: email, name: cleanName(t.name) || cleanName(email.split('@')[0]) || 'Google 使用者' }
}
function validKey(k) {
  const key = String(k || '')
  return key.length >= 16 && key.length <= 64 && /^[\w-]+$/.test(key)
}
/** 這個人再多一支手機（最多 5 支，舊的先掉） */
function addDevice(m, h) {
  if (hasHash(m, h)) return
  m.hs = [h].concat(m.hs || []).slice(0, MAX_DEVICES - 1)
}
/**
 * 用 Google 登入加入（不用連結碼）：
 * - 擁有者綁定過的信箱 → 這支手機成為擁有者的手機；已經是成員（綁過或核准過）→ 這支手機也算他的
 * - 在預先登記名單 → 直接成為成員（權限照登記的）；都不是 → 排進待核准（管理員看得到信箱）
 */
function joinGoogle(data, props) {
  // 這是公開的請求：驗證要打 Google（慢、算每日 UrlFetch 額度），所以放在鎖外面；先在本機擋掉明顯不對的；
  // 打了 Google 還是失敗的次數太多（有人亂送）就暫停 10 分鐘，不然同步會被卡住、額度會被用光
  if (!validKey(data.newKey)) return { ok: false, error: '登入的格式不對：請把 App 更新到最新版再試' }
  const cache = CacheService.getScriptCache()
  if (Number(cache.get('GFAIL') || 0) >= 30) return { ok: false, locked: true, error: 'Google 登入暫時停用（失敗太多次），請 10 分鐘後再試，或改用店裡通行碼' }
  const pre = googleLocalCheck(data.idToken, data.newKey, props.getProperty('GOOGLE_CLIENT_ID'))
  if (pre) return { ok: false, error: pre }
  const v = googleVerify(data.idToken, data.newKey, props.getProperty('GOOGLE_CLIENT_ID'))
  if (v.error) {
    cache.put('GFAIL', String(Number(cache.get('GFAIL') || 0) + 1), 600)
    return { ok: false, error: v.error }
  }
  const lock = LockService.getScriptLock()
  lock.waitLock(20000)
  try {
    const P = props.getProperties()
    // 搬家了：登入驗過才告訴他新網址
    if (P.SYNC_MOVED) return { ok: false, moved: P.SYNC_MOVED, error: '資料已經搬到新的位置' }
    const h = sha(data.newKey)
    const now = Date.now()
    if (P.OWNER_EMAIL && String(P.OWNER_EMAIL).toLowerCase() === v.email) {
      const hs = JSON.parse(P.OWNER_HS || '[]')
      if (hs.indexOf(h) < 0) props.setProperty('OWNER_HS', JSON.stringify([h].concat(hs).slice(0, MAX_DEVICES)))
      return { ok: true, joined: true, role: 'owner' }
    }
    const list = membersOf(P)
    const m = list.filter(function (x) {
      return (x.email || '').toLowerCase() === v.email
    })[0]
    if (m) {
      addDevice(m, h)
      props.setProperty('SYNC_MEMBERS', JSON.stringify(list))
      return { ok: true, joined: true, role: m.role }
    }
    const allow = JSON.parse(P.GOOGLE_ALLOW || '[]')
    const a = allow.filter(function (x) {
      return x.email === v.email
    })[0]
    if (a) {
      list.push({ id: Utilities.getUuid(), name: a.name || v.name, role: a.role, h: h, email: v.email, at: now, seen: 0 })
      props.setProperty('SYNC_MEMBERS', JSON.stringify(list))
      props.setProperty(
        'GOOGLE_ALLOW',
        JSON.stringify(
          allow.filter(function (x) {
            return x.email !== v.email
          }),
        ),
      )
      return { ok: true, joined: true, role: a.role }
    }
    let pend = pendingList(P.JOIN_PENDING)
    if (
      pend.some(function (x) {
        return x.h === h
      })
    )
      return { ok: true, pending: true }
    // 同一個信箱只留最新一筆（每次登入手機都會產生新的連結碼，不然一個帳號就能把 20 個名額塞滿）
    pend = pend.filter(function (x) {
      return (x.email || '') !== v.email
    })
    if (pend.length >= JOIN_MAX_PENDING) return { ok: false, error: '等核准的人太多了：請管理員先處理，再登入' }
    pend.push({ id: Utilities.getUuid(), name: v.name, email: v.email, h: h, at: now })
    props.setProperty('JOIN_PENDING', JSON.stringify(pend))
    return { ok: true, pending: true }
  } finally {
    lock.releaseLock()
  }
}
/** 已經加入的人綁定自己的 Google（擁有者綁了，之後換手機用 Google 登入就能接回來）：登入的 nonce 要等於這支手機現在的連結碼 */
function bindGoogle(data, props, who) {
  const lock = LockService.getScriptLock()
  lock.waitLock(20000)
  try {
    const P = props.getProperties()
    const v = googleVerify(data.idToken, data.key, P.GOOGLE_CLIENT_ID)
    if (v.error) return { ok: false, error: v.error }
    if (who.role === 'owner') {
      props.setProperty('OWNER_EMAIL', v.email)
      return { ok: true, email: v.email }
    }
    const list = membersOf(P)
    const taken = list.some(function (x) {
      return (x.email || '').toLowerCase() === v.email && x.id !== who.id
    })
    if ((P.OWNER_EMAIL || '').toLowerCase() === v.email || taken) return { ok: false, error: '這個 Google 帳號已經被別人綁定了' }
    const m = list.filter(function (x) {
      return x.id === who.id
    })[0]
    if (!m) return { ok: false, error: '找不到你（可能已經被移除）' }
    m.email = v.email
    props.setProperty('SYNC_MEMBERS', JSON.stringify(list))
    return { ok: true, email: v.email }
  } finally {
    lock.releaseLock()
  }
}
/** 同事申請加入：通行碼對、而且沒被暫停，才排進待核准；密碼（連結碼）是他自己的手機產生的，這裡只存雜湊值 */
function joinRequest(data, props, P0) {
  const LOCKED = { ok: false, locked: true, error: '通行碼猜錯太多次，暫時不能申請：請 10 分鐘後再試，或請管理員直接邀請你' }
  const OFF = { ok: false, error: '店裡的通行碼還沒開啟：請找擁有者或管理員' }
  // 這是公開的請求：通行碼沒開、或暫停中，先用 doPost 已經讀好的設定直接擋掉，不要為了它去搶全域鎖（不然有人一直送請求，會拖慢大家的同步）
  if (!normCode(P0.JOIN_CODE)) return OFF
  const f0 = JSON.parse(P0.JOIN_FAILS || '{}')
  if (f0.until && Date.now() < f0.until) return LOCKED
  const lock = LockService.getScriptLock()
  lock.waitLock(20000)
  try {
    const P = props.getProperties()
    const code = normCode(P.JOIN_CODE)
    if (!code) return OFF
    const now = Date.now()
    const f = JSON.parse(P.JOIN_FAILS || '{}')
    if (f.until && now < f.until) return LOCKED
    if (normCode(data.code) !== code) {
      const n = (f.t && now - f.t < JOIN_FAIL_WINDOW ? f.n || 0 : 0) + 1
      props.setProperty('JOIN_FAILS', JSON.stringify(n >= JOIN_MAX_FAILS ? { n: 0, t: now, until: now + JOIN_LOCK_MS } : { n: n, t: now }))
      return { ok: false, error: '通行碼不對' }
    }
    if (P.JOIN_FAILS) props.deleteProperty('JOIN_FAILS')
    // 搬家了：通行碼對的人才告訴他新網址（網址不用白白給外人）
    if (P.SYNC_MOVED) return { ok: false, moved: P.SYNC_MOVED, error: '資料已經搬到新的位置' }
    const name = cleanName(data.name)
    const key = String(data.newKey || '')
    if (!name) return { ok: false, error: '請填名字' }
    if (key.length < 16 || key.length > 64 || !/^[\w-]+$/.test(key)) return { ok: false, error: '申請的格式不對：請把 App 更新到最新版再試' }
    const h = sha(key)
    // 已經核准了（例如核准後又按了一次）：不用重新排隊
    if (
      membersOf(P).some(function (x) {
        return hasHash(x, h)
      })
    )
      return { ok: true, joined: true }
    const pend = pendingList(P.JOIN_PENDING)
    if (
      pend.some(function (x) {
        return x.h === h
      })
    )
      return { ok: true, pending: true }
    if (pend.length >= JOIN_MAX_PENDING) return { ok: false, error: '等核准的人太多了：請管理員先處理，再申請' }
    pend.push({ id: Utilities.getUuid(), name: name, h: h, at: now })
    props.setProperty('JOIN_PENDING', JSON.stringify(pend))
    return { ok: true, pending: true }
  } finally {
    lock.releaseLock()
  }
}
/** 這個人（成員）有沒有這支手機的連結碼雜湊：h＝最早那支、hs＝後來用 Google 登入的手機 */
function hasHash(m, h) {
  return m.h === h || (m.hs || []).indexOf(h) >= 0
}
function allHashes(m) {
  return [m.h].concat(m.hs || [])
}
function membersOf(P) {
  return JSON.parse(P.SYNC_MEMBERS || '[]')
}
/** 這個連結碼是誰：擁有者（第一台）或共用名單裡的人 */
function whoIs(P, key) {
  if (!key) return null
  const h = sha(key)
  if (key === P.SYNC_KEY || JSON.parse(P.OWNER_HS || '[]').indexOf(h) >= 0) return { id: 'owner', name: P.SYNC_OWNER_NAME || '擁有者', role: 'owner' }
  const m = membersOf(P).filter(function (x) {
    return hasHash(x, h)
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
    // 指令碼屬性一個值最多 9 KB：太大就不寫（回 false），由呼叫的地方告訴管理員先移除不用的人
    const save = function () {
      const text = JSON.stringify(list)
      if (encodeURIComponent(text).replace(/%[0-9A-F]{2}/g, 'x').length > 8500) return false
      props.setProperty('SYNC_MEMBERS', text)
      return true
    }
    const FULL = { ok: false, error: '成員太多，雲端的設定放不下：請先移除不用的人' }
    const okRole = function (r) {
      return r === 'manager' || r === 'editor' || r === 'viewer'
    }
    if (data.action === 'members') {
      const people = [{ id: 'owner', name: P.SYNC_OWNER_NAME || '擁有者', role: 'owner', seen: Number(P.SYNC_OWNER_SEEN || 0) }].concat(
        list.map(function (m) {
          return { id: m.id, name: m.name, role: m.role, at: m.at, seen: m.seen || 0, email: m.email || '' }
        }),
      )
      return { ok: true, people: people }
    }
    if (data.action === 'joinInfo') {
      return {
        ok: true,
        code: props.getProperty('JOIN_CODE') || '',
        pending: pendingList(props.getProperty('JOIN_PENDING')).map(function (x) {
          return { id: x.id, name: x.name, at: x.at, email: x.email || '' }
        }),
        allow: JSON.parse(props.getProperty('GOOGLE_ALLOW') || '[]'),
        googleClientId: props.getProperty('GOOGLE_CLIENT_ID') || '',
        ownerEmail: props.getProperty('OWNER_EMAIL') || '',
      }
    }
    if (data.action === 'googleSetup') {
      // 只有擁有者能設定 Google 登入；Client ID 不是密碼，但設錯大家就登入不了
      if (who.role !== 'owner') return { ok: false, error: '只有擁有者可以設定 Google 登入' }
      const cid = String(data.clientId || '').trim()
      if (!cid) props.deleteProperty('GOOGLE_CLIENT_ID')
      else if (!/^[0-9]+-[a-z0-9]+\.apps\.googleusercontent\.com$/.test(cid)) return { ok: false, error: '這不是 Google 登入用的 Client ID（長得像 123456-abc.apps.googleusercontent.com）' }
      else props.setProperty('GOOGLE_CLIENT_ID', cid)
      return { ok: true, googleClientId: cid }
    }
    if (data.action === 'googleAllow' || data.action === 'googleUnallow') {
      const email = String(data.email || '').trim().toLowerCase()
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || email.length > 100) return { ok: false, error: '信箱格式不對' }
      const allow = JSON.parse(props.getProperty('GOOGLE_ALLOW') || '[]').filter(function (x) {
        return x.email !== email
      })
      if (data.action === 'googleAllow') {
        const role = okRole(data.role) ? data.role : 'viewer'
        if (role === 'manager' && who.role !== 'owner') return { ok: false, error: '只有擁有者可以登記管理員' }
        if (allow.length >= 50) return { ok: false, error: '預先登記最多 50 個信箱' }
        allow.push({ email: email, name: cleanName(data.name), role: role, at: Date.now() })
      }
      props.setProperty('GOOGLE_ALLOW', JSON.stringify(allow))
      return { ok: true, allow: allow }
    }
    if (data.action === 'joinCodeSet') {
      if (data.mode === 'off') props.deleteProperty('JOIN_CODE')
      else props.setProperty('JOIN_CODE', newJoinCode())
      props.deleteProperty('JOIN_FAILS')
      return { ok: true, code: props.getProperty('JOIN_CODE') || '' }
    }
    if (data.action === 'joinApprove' || data.action === 'joinReject') {
      const pend = pendingList(props.getProperty('JOIN_PENDING'))
      const i = pend.map(function (x) { return x.id }).indexOf(data.id)
      if (i < 0) return { ok: false, error: '這筆申請不見了（可能已經被別人處理，或超過 7 天過期了）' }
      if (data.action === 'joinApprove') {
        const role = okRole(data.role) ? data.role : 'viewer'
        if (role === 'manager' && who.role !== 'owner') return { ok: false, error: '只有擁有者可以核准成管理員' }
        const p = pend[i]
        list.push({ id: p.id, name: cleanName(data.name) || p.name, role: role, h: p.h, email: p.email || '', at: Date.now(), seen: 0 })
        if (!save()) return FULL
      }
      pend.splice(i, 1)
      props.setProperty('JOIN_PENDING', JSON.stringify(pend))
      return { ok: true }
    }
    if (data.action === 'invite') {
      const name = String(data.name || '').trim().slice(0, 30)
      if (!name) return { ok: false, error: '請填名字' }
      if (!okRole(data.role)) return { ok: false, error: '權限不對' }
      if (data.role === 'manager' && who.role !== 'owner') return { ok: false, error: '只有擁有者可以邀請管理員' }
      if (!data.newKey || String(data.newKey).length < 16) return { ok: false, error: '連結碼太短' }
      const id = Utilities.getUuid()
      list.push({ id: id, name: name, role: data.role, h: sha(data.newKey), at: Date.now(), seen: 0 })
      if (!save()) return FULL
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
      if ((data.role === 'manager' || m.role === 'manager') && who.role !== 'owner') return { ok: false, error: '只有擁有者可以新增或改管理員' }
      m.role = data.role
    } else if (data.action === 'rename') {
      m.name = String(data.name || '').trim().slice(0, 30) || m.name
    } else if (data.action === 'reissue' || data.action === 'remove') {
      // 舊的連結碼列入黑名單：那台下次連線會自動清除資料
      const revoked = JSON.parse(props.getProperty('SYNC_REVOKED') || '[]').concat(allHashes(m))
      props.setProperty('SYNC_REVOKED', JSON.stringify(revoked.slice(-100)))
      if (data.action === 'remove') list.splice(list.indexOf(m), 1)
      else {
        if (!data.newKey || String(data.newKey).length < 16) return { ok: false, error: '連結碼太短' }
        m.h = sha(data.newKey)
        m.hs = []
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
