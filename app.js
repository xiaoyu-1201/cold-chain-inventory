/**
 * 拍照盤點（冷凍材料行）
 * 拍貨架 → Gemini 視覺模型找出每個商品並框起來 → 原圖對照、＋／－ 修正 → 存在手機、匯出 CSV。
 * 沒有後端：API Key 只存在這支手機（localStorage），照片只送到 Google Gemini 分析。
 */

const API = 'https://generativelanguage.googleapis.com/v1beta'
const LS = { key: 'inventory:apiKey', model: 'inventory:model', catalog: 'inventory:catalog' }
const MAX_SIDE = 1600 // 照片先縮到長邊 1600px 再上傳：夠看清楚，又快

/** 店內品項清單（預設值；可以在設定裡改）：給 AI 統一名稱、給修正時選 */
const DEFAULT_CATALOG = `壓縮機（全密閉、半密閉；看銘牌型號）
冷凝器（散熱器；有外箱、裸露型「無穿衫」）
蒸發器（冷排、冷風機；毛細管型、膨脹閥型）
膨脹閥（TE 系列）
膨脹閥閥芯
電磁閥（EVR 系列）
電磁閥線圈
乾燥過濾器（DML 082、083、084、163、164；尾巴 S＝焊接）
視液鏡（SGI、SGN）
手閥（球閥 GBC；2分～7分）
壓力開關（KP 15、KP 1、KP 5）
溫控器（電子式，附感溫棒）
液氣分離器
油分離器
儲液器
銅管（2分、3分、4分、5分、6分、7分、1吋1分、1吋3分、1吋5分）
保溫管（依銅管分數）
毛細管
冷媒（R22、R404A、R507A、R134a、R410A、R32、R417A；鋼瓶）
冷凍油（POE、礦物油；4 公升＝一加侖）
除霜電熱管
風扇馬達
風扇調速器
鰭片清洗劑
銅管接頭（彎頭、三通、直接頭）
喇叭口螺帽`

/** 每一種品項一個顏色（框和清單同色） */
const COLORS = ['#0a84ff', '#30a46c', '#e5484d', '#f76b15', '#8e4ec6', '#0091ff', '#d6409f', '#12a594', '#ad7f58', '#5b5bd6', '#c2a300', '#e54666']

// ───────────────────────── 小工具 ─────────────────────────
const $app = document.getElementById('app')
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c])
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36)
const norm = (s) => String(s ?? '').replace(/\s+/g, '').toLowerCase()
const fmtTime = (t) => new Date(t).toLocaleString('zh-TW', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false })
const ls = {
  get: (k, d = '') => {
    try {
      return localStorage.getItem(k) ?? d
    } catch {
      return d
    }
  },
  set: (k, v) => {
    try {
      localStorage.setItem(k, v)
    } catch {
      /* 無痕模式存不進去就算了 */
    }
  },
}
const catalogLines = () => (ls.get(LS.catalog) || DEFAULT_CATALOG).split('\n').map((l) => l.trim()).filter(Boolean)
const catalogNames = () => catalogLines().map((l) => l.replace(/（.*$/, '').trim())

let toastTimer = 0
function toast(msg) {
  document.querySelector('.toast')?.remove()
  const el = document.createElement('div')
  el.className = 'toast'
  el.setAttribute('role', 'status')
  el.textContent = msg
  document.body.appendChild(el)
  clearTimeout(toastTimer)
  toastTimer = setTimeout(() => el.remove(), 2800)
}

// ───────────────────────── 存檔（IndexedDB，照片也存在手機） ─────────────────────────
const db = (() => {
  let p
  const open = () =>
    (p ??= new Promise((resolve, reject) => {
      const req = indexedDB.open('inventory', 1)
      req.onupgradeneeded = () => req.result.createObjectStore('sessions', { keyPath: 'id' })
      req.onsuccess = () => resolve(req.result)
      req.onerror = () => reject(req.error)
    }))
  const tx = async (mode, fn) => {
    const d = await open()
    return new Promise((resolve, reject) => {
      const t = d.transaction('sessions', mode)
      const r = fn(t.objectStore('sessions'))
      t.oncomplete = () => resolve(r?.result)
      t.onerror = () => reject(t.error)
    })
  }
  return {
    all: async () => ((await tx('readonly', (s) => s.getAll())) ?? []).sort((a, b) => b.createdAt - a.createdAt),
    get: (id) => tx('readonly', (s) => s.get(id)),
    put: (session) => tx('readwrite', (s) => s.put(session)),
    del: (id) => tx('readwrite', (s) => s.delete(id)),
    clear: () => tx('readwrite', (s) => s.clear()),
  }
})()

// ───────────────────────── 照片：縮圖 ─────────────────────────
async function prepareImage(file) {
  // 舊瀏覽器不支援 imageOrientation 參數就退回預設（新版預設也會依 EXIF 轉正）
  const bmp = await createImageBitmap(file, { imageOrientation: 'from-image' }).catch(() => createImageBitmap(file))
  const scale = Math.min(1, MAX_SIDE / Math.max(bmp.width, bmp.height))
  const w = Math.round(bmp.width * scale)
  const h = Math.round(bmp.height * scale)
  const canvas = document.createElement('canvas')
  canvas.width = w
  canvas.height = h
  canvas.getContext('2d').drawImage(bmp, 0, 0, w, h)
  bmp.close?.()
  const blob = await new Promise((r) => canvas.toBlob(r, 'image/jpeg', 0.85))
  return { id: uid(), blob, w, h, objects: [], status: 'pending' }
}
const blobToBase64 = (blob) =>
  new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = () => resolve(String(r.result).split(',')[1])
    r.onerror = () => reject(r.error)
    r.readAsDataURL(blob)
  })
const urls = new Map()
const urlOf = (photo) => {
  if (!urls.has(photo.id)) urls.set(photo.id, URL.createObjectURL(photo.blob))
  return urls.get(photo.id)
}

// ───────────────────────── Gemini ─────────────────────────
class ApiError extends Error {
  constructor(message, status, detail = '') {
    super(message)
    this.status = status
    /** 原始錯誤（狀態碼＋Google 的訊息），顯示在小字，方便截圖回報 */
    this.detail = detail
    /** Google 建議等幾秒再試（429 會附） */
    this.retryAfter = 0
  }
}
function friendly(status, msg = '') {
  if (status === 400 && /api key/i.test(msg)) return 'API Key 不對：請到「設定」重新貼上。'
  if (status === 400) return `AI 看不懂這次的請求（${msg.slice(0, 80)}）`
  if (status === 403) return '這個 API Key 沒有權限用 Gemini，請確認是在 Google AI Studio 建立的金鑰。'
  if (status === 404) return '這個模型已經下架或不能用，已自動改用其他模型，請再試一次。'
  if (status === 429) return '免費額度一分鐘內用太多次了，等 1 分鐘再試。'
  if (status === 503) return 'Google 的 AI 現在太多人用（免費版尖峰常見），已經自動重試和換模型，還是不行；等幾分鐘再按「再試一次」。'
  if (status >= 500) return 'Google 那邊暫時出問題，已經自動重試；等一下再按「再試一次」。'
  return msg || '連線失敗，請確認有網路。'
}
async function call(path, opts = {}) {
  const key = ls.get(LS.key)
  let res
  try {
    res = await fetch(`${API}/${path}${path.includes('?') ? '&' : '?'}key=${encodeURIComponent(key)}`, opts)
  } catch {
    throw new ApiError('沒有網路，或連不到 Google。', 0, 'network error')
  }
  const data = await res.json().catch(() => ({}))
  if (!res.ok) {
    const msg = data?.error?.message || ''
    const err = new ApiError(friendly(res.status, msg), res.status, `${res.status} ${data?.error?.status || ''} ${msg}`.trim().slice(0, 200))
    // 429 會告訴你要等多久（例如 "17s"）
    const retry = (data?.error?.details || []).find((d) => d.retryDelay)?.retryDelay
    if (retry) err.retryAfter = Math.min(30, parseFloat(retry) || 0)
    throw err
  }
  return data
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
/** 這些狀況值得自動重試：太多人用、伺服器暫時錯誤、額度每分鐘上限 */
const retryable = (e) => e instanceof ApiError && [429, 500, 502, 503, 504].includes(e.status)

/** 自動挑模型：可用、會看圖、名字有 flash；穩定版優先、版本新的優先、lite 排後面 */
function rankModels(models) {
  const score = (name) => {
    const v = parseFloat((name.match(/gemini-(\d+(?:\.\d+)?)/) || [])[1] || '0')
    const unstable = /preview|exp/.test(name) ? 1 : 0
    const lite = /lite/.test(name) ? 1 : 0
    const latest = /latest/.test(name) ? 0.05 : 0
    return v * 10 - unstable * 25 - lite * 6 + latest
  }
  return models
    .filter((m) => (m.supportedGenerationMethods || []).includes('generateContent'))
    .map((m) => m.name.replace(/^models\//, ''))
    .filter((n) => /flash/.test(n) && !/image|tts|audio|live|embedding|thinking|8b/.test(n))
    .sort((a, b) => score(b) - score(a))
}
let modelCache = null
async function fetchModels() {
  const data = await call('models?pageSize=200')
  modelCache = rankModels(data.models || [])
  return modelCache
}
async function currentModel(force = false) {
  const saved = ls.get(LS.model)
  if (saved && !force) return saved
  const list = await fetchModels()
  if (!list.length) throw new ApiError('這個 API Key 找不到能看圖的 Gemini 模型。', 0)
  ls.set(LS.model, list[0])
  return list[0]
}

const SCHEMA = {
  type: 'OBJECT',
  properties: {
    objects: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          label: { type: 'STRING', description: '品名，優先用店內品項清單的寫法' },
          brand: { type: 'STRING', description: '品牌（看標籤或外盒）；看不出來就空字串' },
          model: { type: 'STRING', description: '型號（例如 DML 083、KP 15）；看不出來就空字串' },
          spec: { type: 'STRING', description: '尺寸／容量／規格（例如 3分、10.9kg、4L）；看不出來就空字串' },
          box_2d: { type: 'ARRAY', items: { type: 'INTEGER' }, description: '[ymin, xmin, ymax, xmax]，0～1000' },
          confidence: { type: 'NUMBER', description: '0～1，有多確定' },
        },
        required: ['label', 'brand', 'model', 'spec', 'box_2d', 'confidence'],
      },
    },
    note: { type: 'STRING', description: '看不清楚、被擋住、需要人工確認的地方；沒有就空字串' },
  },
  required: ['objects', 'note'],
}

function prompt() {
  return `你是冷凍空調材料行的盤點助手。請找出照片裡每一個「商品」，每一個各給一個框。
規則：
1. 名稱優先用下面「店內品項清單」的寫法；清單沒有就用最具體的中文名稱（例如「乾燥過濾器」，不要只寫「零件」）。
2. 標籤看得到就分開填：brand＝品牌（例如 Danfoss）、model＝型號（例如 DML 083、KP 15）、spec＝尺寸／容量／規格（例如 3分、10.9kg、4L、R404A）。看不出來就空字串，不要猜。
3. 同一種東西有幾個就給幾個框，不要合併成一個；品牌、型號或尺寸不同就算不同的東西。被擋住一半以上的也要算，但 confidence 給低一點。
4. 貨架、標價牌、手、紙箱外的雜物不算商品。
5. box_2d 用 [ymin, xmin, ymax, xmax]，是相對整張照片的 0～1000。
店內品項清單：
${catalogLines().join('\n')}`
}

/** 用某個模型送一次辨識 */
async function generate(model, image, thinkingOff = true) {
  const body = {
    contents: [{ role: 'user', parts: [{ inline_data: { mime_type: 'image/jpeg', data: image } }, { text: prompt() }] }],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'application/json',
      responseSchema: SCHEMA,
      // 2.5 Flash 預設會先「想」很久；關掉思考，速度快很多
      ...(thinkingOff && /2\.5-flash/.test(model) ? { thinkingConfig: { thinkingBudget: 0 } } : {}),
    },
  }
  const data = await call(`models/${model}:generateContent`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
  const cand = data?.candidates?.[0]
  const text = cand?.content?.parts?.map((p) => p.text || '').join('') || ''
  if (!text) throw new ApiError(cand?.finishReason === 'SAFETY' ? 'AI 拒絕分析這張照片，請換一張。' : 'AI 沒有回傳結果，請再試一次。', 0, `empty response ${cand?.finishReason || ''}`)
  let parsed
  try {
    parsed = JSON.parse(text)
  } catch {
    throw new ApiError('AI 回傳的格式壞掉了，請再試一次。', 0, `bad json: ${text.slice(0, 120)}`)
  }
  return {
    objects: (parsed.objects || [])
      .filter((o) => Array.isArray(o.box_2d) && o.box_2d.length === 4 && o.label)
      .map((o) => ({
        label: String(o.label).trim(),
        brand: String(o.brand || '').trim(),
        model: String(o.model || '').trim(),
        spec: String(o.spec || '').trim(),
        box: o.box_2d.map((n) => Math.min(1000, Math.max(0, Number(n) || 0))),
        confidence: Math.min(1, Math.max(0, Number(o.confidence) || 0)),
      })),
    note: String(parsed.note || ''),
    model,
  }
}

/**
 * 分析一張照片：同一個模型遇到「太多人用」會等一下重試（2、4、8 秒）；
 * 還是不行就換下一個模型（例如 Flash → Flash-Lite）；模型下架（404）就重新抓清單。
 * onStatus 用來在畫面上顯示目前在做什麼。
 */
async function analyze(photo, onStatus = () => {}) {
  const image = await blobToBase64(photo.blob)
  const first = await currentModel()
  const list = modelCache ?? (await fetchModels().catch(() => [first]))
  const models = [first, ...list.filter((m) => m !== first)].slice(0, 3)
  let lastErr
  for (const [mi, model] of models.entries()) {
    if (mi > 0) onStatus(`改用 ${model} 再試`)
    let thinkingOff = true
    for (let attempt = 0; attempt < 4; attempt++) {
      try {
        return await generate(model, image, thinkingOff)
      } catch (e) {
        lastErr = e
        if (state.cancel) throw e
        // 有的模型不接受「關掉思考」的參數：拿掉再試一次
        if (e instanceof ApiError && e.status === 400 && /thinking/i.test(e.detail) && thinkingOff) {
          thinkingOff = false
          continue
        }
        if (e instanceof ApiError && e.status === 404) {
          ls.set(LS.model, '')
          break
        }
        if (!retryable(e) || attempt === 3) break
        const wait = e.retryAfter ? e.retryAfter * 1000 : 2000 * 2 ** attempt
        onStatus(`Google 忙線（${e.status}），${Math.round(wait / 1000)} 秒後自動重試（第 ${attempt + 2} 次）`)
        await sleep(wait)
      }
    }
    // 只有「忙線、暫時錯誤、模型下架」才換模型；Key 錯、照片問題換模型也沒用
    if (!(retryable(lastErr) || lastErr?.status === 404)) break
  }
  throw lastErr
}

// ───────────────────────── 盤點結果：把框合併成品項 ─────────────────────────
/** 品名、品牌、型號、規格四個都一樣才算同一種 */
const FIELDS = ['label', 'brand', 'model', 'spec']
const keyOf = (o) => FIELDS.map((f) => norm(o[f])).join('|')
/** 清單第二行：品牌・型號・規格 */
const detailOf = (o) => [o.brand, o.model, o.spec].filter(Boolean).join('・')

function groupsOf(session) {
  const map = new Map()
  session.photos.forEach((photo, pi) =>
    photo.objects.forEach((o, oi) => {
      const key = keyOf(o)
      if (!map.has(key)) map.set(key, { key, label: o.label, brand: o.brand || '', model: o.model || '', spec: o.spec || '', boxes: 0, conf: 0, edited: false, refs: [] })
      const g = map.get(key)
      g.boxes += 1
      g.conf += o.confidence
      g.edited ||= !!o.edited
      g.refs.push({ pi, oi })
    }),
  )
  // 固定排序（品名 → 尺寸 → 品牌 → 型號），修改後編號和顏色不會亂跳
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  const groups = [...map.values()].map((g) => ({ ...g, conf: g.boxes ? g.conf / g.boxes : 1 })).sort((a, b) => cmp(a.label, b.label) || cmp(a.spec, b.spec) || cmp(a.brand, b.brand) || cmp(a.model, b.model))
  for (const m of session.manual || []) groups.push({ key: `manual:${m.id}`, label: m.label, brand: m.brand || '', model: m.model || '', spec: m.spec || '', boxes: 0, conf: 1, refs: [], manual: m })
  return groups.map((g, i) => ({ ...g, color: COLORS[i % COLORS.length], count: g.manual ? g.manual.count : (session.counts?.[g.key] ?? g.boxes) }))
}

/** 改數量（手動新增的存在品項上；AI 的存成覆寫值） */
function setCount(session, g, n) {
  const count = Math.max(0, n)
  if (g.manual) g.manual.count = count
  else (session.counts ??= {})[g.key] = count
}

/**
 * 把一些框改成另一個品項（整組改或只改一個框）。
 * 如果原本有手動改過的數量：搬過去；跟別的品項合併時，數量相加。
 */
function moveObjects(session, refs, fields) {
  const counts = (session.counts ??= {})
  const objs = refs.map((r) => session.photos[r.pi].objects[r.oi])
  const oldKey = keyOf(objs[0])
  const oldGroup = groupsOf(session).find((x) => x.key === oldKey)
  objs.forEach((o) => Object.assign(o, fields, { edited: true }))
  const newKey = keyOf(objs[0])
  if (newKey === oldKey) return
  const moved = objs.length
  const whole = oldGroup && moved === oldGroup.boxes
  if (counts[oldKey] !== undefined) {
    const carry = whole ? counts[oldKey] : Math.min(moved, counts[oldKey])
    if (whole) delete counts[oldKey]
    else counts[oldKey] = Math.max(0, counts[oldKey] - moved)
    const target = groupsOf(session).find((x) => x.key === newKey)
    counts[newKey] = (counts[newKey] ?? (target ? target.boxes - moved : 0)) + carry
  } else if (counts[newKey] !== undefined) counts[newKey] += moved
}

/** 刪掉一些框（照片同一張要從後面刪，索引才不會亂） */
function removeObjects(session, refs) {
  const counts = session.counts ?? {}
  const key = keyOf(session.photos[refs[0].pi].objects[refs[0].oi])
  ;[...refs].sort((a, b) => b.oi - a.oi).forEach((r) => session.photos[r.pi].objects.splice(r.oi, 1))
  if (counts[key] !== undefined) {
    const left = groupsOf(session).find((x) => x.key === key)
    if (!left) delete counts[key]
    else counts[key] = Math.max(0, counts[key] - refs.length)
  }
}
const totalQty = (session) => groupsOf(session).reduce((n, g) => n + (Number(g.count) || 0), 0)

// ───────────────────────── 畫面狀態 ─────────────────────────
const state = { view: 'home', session: null, photoIndex: 0, focus: null, busy: false, cancel: false, progress: null }

function go(view, extra = {}) {
  Object.assign(state, { view, focus: null }, extra)
  render()
  window.scrollTo({ top: 0 })
}
async function save() {
  if (state.session) await db.put(state.session)
}

// ───────────────────────── 畫面 ─────────────────────────
const chev = '<svg class="chev" width="10" height="17" viewBox="0 0 10 17" aria-hidden="true"><path d="M1.5 1.5 8 8.5l-6.5 7" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>'
const backBtn = (to = 'home', label = '盤點') => `<button class="back" data-go="${to}"><svg width="12" height="20" viewBox="0 0 12 20" aria-hidden="true"><path d="M10 2 2 10l8 8" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>${esc(label)}</button>`

async function viewHome() {
  const sessions = await db.all()
  const hasKey = !!ls.get(LS.key)
  return `
  <main class="app">
    <div class="nav"><span></span><button class="icon-btn" data-go="settings" aria-label="設定"><svg width="22" height="22" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 15.5A3.5 3.5 0 1 1 12 8.5a3.5 3.5 0 0 1 0 7Zm7.43-2.53a7.8 7.8 0 0 0 0-1.94l2.11-1.65a.5.5 0 0 0 .12-.64l-2-3.46a.5.5 0 0 0-.6-.22l-2.49 1a7.6 7.6 0 0 0-1.68-.98l-.38-2.65A.5.5 0 0 0 14 2h-4a.5.5 0 0 0-.5.42l-.37 2.65c-.61.25-1.17.58-1.69.98l-2.49-1a.5.5 0 0 0-.6.22l-2 3.46a.5.5 0 0 0 .12.64l2.11 1.65a7.8 7.8 0 0 0 0 1.94l-2.11 1.65a.5.5 0 0 0-.12.64l2 3.46a.5.5 0 0 0 .6.22l2.49-1c.52.4 1.08.73 1.69.98l.37 2.65c.04.24.25.42.5.42h4c.25 0 .46-.18.5-.42l.37-2.65c.61-.25 1.17-.58 1.68-.98l2.49 1a.5.5 0 0 0 .6-.22l2-3.46a.5.5 0 0 0-.12-.64l-2.1-1.65Z"/></svg></button></div>
    <h1 class="large-title">拍照盤點</h1>
    <p class="subtitle">拍貨架，AI 數品項；跟原圖對照，再用 ＋／－ 修正。</p>
    ${
      hasKey
        ? ''
        : `<div class="hint-card stack"><div><b>第一次用：</b>先到「設定」貼上你的免費 Gemini API Key（只會存在這支手機）。</div><button class="btn small" data-go="settings">去設定</button></div>`
    }
    <p class="section-title">盤點紀錄</p>
    ${
      sessions.length
        ? `<div class="group">${sessions
            .map((s) => {
              const n = groupsOf(s).length
              return `<button class="row" data-open="${s.id}"><img src="${s.photos[0] ? urlOf(s.photos[0]) : ''}" alt="" style="width:48px;height:48px;border-radius:10px;object-fit:cover;background:var(--card-2)"><span class="grow"><span class="title">${esc(s.place || '未命名位置')}</span><br><span class="meta">${fmtTime(s.createdAt)}・${s.photos.length} 張照片・${n} 種、共 ${totalQty(s)} 件</span></span>${chev}</button>`
            })
            .join('')}</div>`
        : `<div class="empty"><div class="big">📦</div><p>還沒有盤點紀錄。<br>按下面「新盤點」，拍一層貨架試試看。</p></div>`
    }
  </main>
  <div class="toolbar"><div class="inner"><button class="btn" data-action="new" ${hasKey ? '' : 'disabled'}>＋ 新盤點</button></div></div>`
}

function viewCapture() {
  const s = state.session
  return `
  <main class="app">
    <div class="nav">${backBtn('home', '盤點')}</div>
    <h1 class="large-title">新盤點</h1>
    <div class="group"><label class="row"><span style="width:56px" class="muted">位置</span><input class="inline" id="place" placeholder="例：A 貨架第 2 層" value="${esc(s.place)}" autocomplete="off"></label></div>
    <p class="section-title">照片（${s.photos.length}）</p>
    ${
      s.photos.length
        ? `<div class="photo-strip">${s.photos.map((p, i) => `<button data-remove-photo="${i}" aria-label="移除第 ${i + 1} 張"><img src="${urlOf(p)}" alt=""></button>`).join('')}</div><p class="footnote">點縮圖可以移除那張。</p>`
        : ''
    }
    <div class="row-actions" style="margin-top:12px">
      <label class="btn secondary" style="flex:1">📷 拍照<input type="file" accept="image/*" capture="environment" id="cam" class="sr-only"></label>
      <label class="btn secondary" style="flex:1">🖼 從相簿選<input type="file" accept="image/*" multiple id="pick" class="sr-only"></label>
    </div>
    <p class="section-title">怎麼拍比較準</p>
    <div class="group">
      <div class="row"><span>① 一次拍一層貨架，正面拍，不要太斜</span></div>
      <div class="row"><span>② 光線夠、不要反光；標籤朝外最好</span></div>
      <div class="row"><span>③ 疊在一起的、被擋住的，結果出來再用 ＋／－ 修正</span></div>
    </div>
  </main>
  <div class="toolbar"><div class="inner"><button class="btn" data-action="analyze" ${s.photos.length ? '' : 'disabled'}>開始辨識${s.photos.length ? `（${s.photos.length} 張）` : ''}</button></div></div>`
}

function viewAnalyzing() {
  const p = state.progress || { done: 0, total: 0, started: Date.now() }
  return `
  <main class="app">
    <div class="progress">
      <div class="spinner" aria-hidden="true"></div>
      <div><b style="color:var(--text)">AI 辨識中… ${p.done + 1 > p.total ? p.total : p.done + 1} / ${p.total}</b><br>每張大約 5～15 秒（看網路和照片內容）</div>
      <div class="muted" id="elapsed">已經 0 秒</div>
      <div class="muted" id="ai-status" style="min-height:1.5em;font-size:15px"></div>
    </div>
  </main>
  <div class="toolbar"><div class="inner"><button class="btn plain" data-action="cancel">取消</button></div></div>`
}

function viewReview() {
  const s = state.session
  const groups = groupsOf(s)
  const photo = s.photos[state.photoIndex]
  const colorOfRef = new Map()
  groups.forEach((g, gi) => g.refs.forEach((r) => colorOfRef.set(`${r.pi}:${r.oi}`, { color: g.color, gi, key: g.key })))
  const boxes = photo
    ? photo.objects
        .map((o, oi) => {
          const ref = colorOfRef.get(`${state.photoIndex}:${oi}`)
          if (!ref) return ''
          const [y1, x1, y2, x2] = o.box
          const on = state.focus === ref.key
          return `<div class="box ${on ? 'on' : ''}" style="--c:${ref.color};top:${y1 / 10}%;left:${x1 / 10}%;height:${(y2 - y1) / 10}%;width:${(x2 - x1) / 10}%" data-focus="${esc(ref.key)}" data-obj="${state.photoIndex}:${oi}"><span class="tag">${ref.gi + 1}</span></div>`
        })
        .join('')
    : ''
  const errors = s.photos.map((p, i) => (p.status !== 'done' ? `<div class="error-card">第 ${i + 1} 張沒辨識成功：${esc(p.error || '還沒辨識（被取消）')} <button class="btn small secondary" data-retry="${i}" style="margin-left:6px">再試一次</button>${p.errorDetail ? `<div style="margin-top:8px;font-size:12px;color:var(--text-2);word-break:break-all">錯誤代碼：${esc(p.errorDetail)}</div>` : ''}</div>` : '')).join('')
  const notes = s.photos.map((p, i) => (p.note ? `<p class="footnote">第 ${i + 1} 張 AI 備註：${esc(p.note)}</p>` : '')).join('')
  return `
  <main class="app">
    <div class="nav">${backBtn('home', '盤點')}<button class="btn small secondary" data-action="add">＋ 手動新增</button></div>
    <h1 class="large-title">${esc(s.place || '盤點結果')}</h1>
    <p class="subtitle">${fmtTime(s.createdAt)}・${groups.length} 種、共 ${totalQty(s)} 件${s.model ? `・${esc(s.model)}` : ''}</p>
    ${errors ? `<div class="stack">${errors}</div>` : ''}
    ${
      photo
        ? `<div class="photo-wrap ${state.focus ? 'focus' : ''}" data-photo><img src="${urlOf(photo)}" alt="第 ${state.photoIndex + 1} 張照片">${boxes}</div>
           ${s.photos.length > 1 ? `<div class="photo-strip">${s.photos.map((p, i) => `<button class="${i === state.photoIndex ? 'on' : ''}" data-photo-index="${i}" aria-label="看第 ${i + 1} 張"><img src="${urlOf(p)}" alt=""></button>`).join('')}</div>` : ''}
           <p class="footnote">${state.focus ? '再點一次框，可以單獨修改那一個（例如尺寸不一樣）。' : '點框或點清單，對照是哪一個；數量不對就按 ＋／－，名稱不對就點 ✎。'}</p>`
        : ''
    }
    ${notes}
    <p class="section-title">品項（${groups.length}）</p>
    ${
      groups.length
        ? `<div class="group">${groups
            .map(
              (g, gi) => `
          <div class="item ${state.focus === g.key ? 'on' : ''}" data-item="${esc(g.key)}">
            <button class="swatch" style="--c:${g.color}" data-focus="${esc(g.key)}" aria-label="在照片上標出 ${esc(g.label)}">${gi + 1}</button>
            <button class="grow edit-btn" data-edit="${esc(g.key)}" aria-label="修改 ${esc(g.label)} 的名稱、品牌、型號、規格">
              <span class="name">${esc(g.label)}</span><span class="pencil" aria-hidden="true">✎</span>${g.manual ? '<span class="badge edit">手動</span>' : g.conf < 0.6 && !g.edited ? '<span class="badge low">請確認</span>' : ''}${g.edited && !g.manual ? '<span class="badge edit">已修正</span>' : ''}
              <br><span class="spec">${esc(detailOf(g) || '點 ✎ 補品牌、型號、尺寸')}${!g.manual && g.count !== g.boxes ? `・照片裡 ${g.boxes} 個` : ''}</span>
            </button>
            <span class="stepper"><button data-step="-1" data-key="${esc(g.key)}" aria-label="減一">−</button><input inputmode="numeric" value="${g.count}" data-count="${esc(g.key)}" aria-label="${esc(g.label)} 數量"><button data-step="1" data-key="${esc(g.key)}" aria-label="加一">＋</button></span>
          </div>`,
            )
            .join('')}</div>`
        : `<div class="empty"><p>這次沒有找到商品。<br>可以重拍，或按右上「手動新增」。</p></div>`
    }
    <div class="row-actions" style="margin-top:22px">
      <button class="btn danger small" data-action="delete-session">刪除這次盤點</button>
    </div>
  </main>
  <div class="toolbar"><div class="inner"><button class="btn secondary" data-action="export">匯出</button><button class="btn" data-go="home">完成</button></div></div>`
}

function viewSettings() {
  const key = ls.get(LS.key)
  const model = ls.get(LS.model)
  return `
  <main class="app">
    <div class="nav">${backBtn('home', '盤點')}</div>
    <h1 class="large-title">設定</h1>
    <p class="section-title">Gemini API Key</p>
    <div class="stack">
      <input class="field" id="apikey" type="password" placeholder="貼上 API Key（AIza 開頭）" value="${esc(key)}" autocomplete="off" spellcheck="false">
      <div class="row-actions"><button class="btn small" data-action="save-key">儲存並測試</button><button class="btn small secondary" data-action="toggle-key">顯示／隱藏</button></div>
      <p class="footnote" style="margin:0">沒有 Key？到 <a href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer" style="color:var(--tint)">Google AI Studio</a> 免費建立。Key 只存在這支手機，照片只會送到 Google Gemini 分析。</p>
    </div>
    <p class="section-title">辨識模型</p>
    <div class="group"><div class="row"><span class="grow"><span class="title">${esc(model || '自動挑選')}</span><br><span class="meta">自動挑能看圖、最新又快的 Flash；被下架會自動換</span></span><button class="btn small secondary" data-action="pick-model">重新挑選</button></div></div>
    <div id="models"></div>
    <p class="section-title">店內品項清單</p>
    <textarea class="field" id="catalog" spellcheck="false">${esc(catalogLines().join('\n'))}</textarea>
    <p class="footnote">一行一種品項，括號裡寫規格或別名。AI 會照這裡的名稱寫，修正時也會跳出來給你選。</p>
    <div class="row-actions" style="margin-top:10px"><button class="btn small" data-action="save-catalog">儲存清單</button><button class="btn small secondary" data-action="reset-catalog">恢復預設</button></div>
    <p class="section-title">資料</p>
    <div class="row-actions"><button class="btn small danger" data-action="clear-all">刪除全部盤點紀錄</button></div>
    <p class="footnote">紀錄（含照片）只存在這支手機的瀏覽器裡；要留底請用「匯出」。</p>
  </main>`
}

async function render() {
  const html = state.view === 'home' ? await viewHome() : state.view === 'capture' ? viewCapture() : state.view === 'analyzing' ? viewAnalyzing() : state.view === 'review' ? viewReview() : viewSettings()
  $app.innerHTML = html
  bindInputs()
}

// ───────────────────────── 底部面板（Sheet）：點旁邊暗處或按 Esc 就關 ─────────────────────────
function sheet(html, onMount) {
  const back = document.createElement('div')
  back.className = 'sheet-backdrop'
  back.innerHTML = `<div class="sheet" role="dialog" aria-modal="true"><div class="grabber"></div>${html}</div>`
  const close = () => {
    back.remove()
    document.removeEventListener('keydown', onKey)
  }
  const onKey = (e) => e.key === 'Escape' && close()
  back.addEventListener('click', (e) => e.target === back && close())
  document.addEventListener('keydown', onKey)
  document.body.appendChild(back)
  onMount?.(back.querySelector('.sheet'), close)
  return close
}

/** 常見品牌（只是輸入時的建議，可以自己打） */
const BRANDS = ['Danfoss', 'Emerson', 'Copeland', 'Sporlan', 'Castel', 'Carel', 'Dixell', 'Bitzer', 'Tecumseh', 'Embraco', 'Panasonic', 'Hitachi', 'Daikin', 'Chemours']

/** 輸入建議：店內品項清單＋常見品牌＋以前盤點打過的字（越常用越前面） */
async function suggestions() {
  const tally = { label: new Map(), brand: new Map(), model: new Map(), spec: new Map() }
  const add = (o) => FIELDS.forEach((f) => o[f] && tally[f].set(o[f], (tally[f].get(o[f]) || 0) + 1))
  for (const s of await db.all()) {
    s.photos.forEach((p) => p.objects.forEach(add))
    ;(s.manual || []).forEach(add)
  }
  state.session?.photos.forEach((p) => p.objects.forEach(add))
  const top = (m) => [...m.entries()].sort((a, b) => b[1] - a[1]).map(([v]) => v)
  return {
    label: [...new Set([...catalogNames(), ...top(tally.label)])],
    brand: [...new Set([...top(tally.brand), ...BRANDS])],
    model: top(tally.model).slice(0, 80),
    spec: [...new Set([...top(tally.spec), '2分', '3分', '4分', '5分', '6分', '7分', '1吋1分', '1吋3分'])],
  }
}

/** 四個欄位的表單（品名、品牌、型號、尺寸／規格） */
function fieldsHtml(v, sug) {
  const row = (f, label, ph) => `
    <label class="field-label" for="f-${f}">${label}</label>
    <input class="field" id="f-${f}" list="dl-${f}" value="${esc(v[f] || '')}" placeholder="${ph}" autocomplete="off">
    <datalist id="dl-${f}">${sug[f].map((n) => `<option value="${esc(n)}">`).join('')}</datalist>`
  return `<div class="form">
    ${row('label', '品名', '例：乾燥過濾器、保溫管')}
    ${row('brand', '品牌', '例：Danfoss（看不出來可以空白）')}
    ${row('model', '型號', '例：DML 083、KP 15')}
    ${row('spec', '尺寸／規格', '例：3分、10.9kg、4L')}
  </div>`
}
const readFields = (el) => Object.fromEntries(FIELDS.map((f) => [f, el.querySelector(`#f-${f}`).value.trim()]))

/** 修改整個品項（這一列的全部框一起改） */
async function editSheet(key) {
  const s = state.session
  const g = groupsOf(s).find((x) => x.key === key)
  if (!g) return
  const sug = await suggestions()
  sheet(
    `<h2 class="sheet-title">修改品項</h2>
     <p class="sheet-sub">${g.manual ? '手動新增的品項' : `照片裡 ${g.boxes} 個框會一起改；只有其中幾個不一樣，請在照片上點那個框`}</p>
     ${fieldsHtml(g, sug)}
     <div class="row-actions" style="margin-top:16px"><button class="btn" style="flex:1" id="e-save">儲存</button><button class="btn danger" id="e-del">刪掉</button></div>`,
    (el, close) => {
      el.querySelector('#e-save').onclick = async () => {
        const v = readFields(el)
        if (!v.label) return toast('品名不能空白')
        if (g.manual) Object.assign(g.manual, v)
        else moveObjects(s, g.refs, v)
        state.focus = null
        await save()
        close()
        render()
      }
      el.querySelector('#e-del').onclick = async () => {
        if (!confirm(`刪掉「${g.label}」${g.manual ? '' : `（${g.boxes} 個框）`}？`)) return
        if (g.manual) s.manual = s.manual.filter((m) => m !== g.manual)
        else removeObjects(s, g.refs)
        state.focus = null
        await save()
        close()
        render()
        toast('已刪掉')
      }
    },
  )
}

/** 只改照片上的某一個框（例如同一排保溫管，有一支尺寸不一樣） */
async function objectSheet(pi, oi) {
  const s = state.session
  const o = s.photos[pi]?.objects[oi]
  if (!o) return
  const sug = await suggestions()
  sheet(
    `<h2 class="sheet-title">修改這一個</h2>
     <p class="sheet-sub">只改照片上這一個框；改完會自動歸到對的品項、數量也會跟著變。</p>
     ${fieldsHtml(o, sug)}
     <div class="row-actions" style="margin-top:16px"><button class="btn" style="flex:1" id="o-save">儲存</button><button class="btn danger" id="o-del">這不是商品，刪掉這個框</button></div>`,
    (el, close) => {
      el.querySelector('#o-save').onclick = async () => {
        const v = readFields(el)
        if (!v.label) return toast('品名不能空白')
        moveObjects(s, [{ pi, oi }], v)
        state.focus = keyOf(o)
        await save()
        close()
        render()
      }
      el.querySelector('#o-del').onclick = async () => {
        removeObjects(s, [{ pi, oi }])
        state.focus = null
        await save()
        close()
        render()
        toast('已刪掉這個框')
      }
    },
  )
}

async function addSheet() {
  const sug = await suggestions()
  sheet(
    `<h2 class="sheet-title">手動新增品項</h2>
     ${fieldsHtml({}, sug)}
     <label class="field-label" for="a-count">數量</label>
     <input class="field" id="a-count" inputmode="numeric" value="1">
     <button class="btn block" id="a-save" style="margin-top:16px">新增</button>`,
    (el, close) => {
      el.querySelector('#f-label').focus()
      el.querySelector('#a-save').onclick = async () => {
        const v = readFields(el)
        if (!v.label) return toast('請先填品名')
        const s = state.session
        ;(s.manual ??= []).push({ id: uid(), ...v, count: Math.max(0, parseInt(el.querySelector('#a-count').value, 10) || 0) })
        await save()
        close()
        render()
      }
    },
  )
}

// ───────────────────────── 匯出 ─────────────────────────
function csvOf(session) {
  const cell = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`
  const rows = [['盤點時間', '位置', '品名', '品牌', '型號', '尺寸／規格', '數量', '照片框數', 'AI 信心', '來源']]
  for (const g of groupsOf(session)) rows.push([fmtTime(session.createdAt), session.place, g.label, g.brand, g.model, g.spec, g.count, g.manual ? '' : g.boxes, g.manual ? '' : g.conf.toFixed(2), g.manual ? '手動' : g.edited ? 'AI（有修正）' : 'AI'])
  return '﻿' + rows.map((r) => r.map(cell).join(',')).join('\r\n') // 加 BOM，Excel 開才不會亂碼
}
const textOf = (session) => [`盤點：${session.place || '未命名位置'}（${fmtTime(session.createdAt)}）`, ...groupsOf(session).map((g) => `・${g.label}${detailOf(g) ? `（${detailOf(g)}）` : ''}：${g.count}`), `共 ${totalQty(session)} 件`].join('\n')

function exportSheet() {
  const s = state.session
  const name = `盤點_${(s.place || '未命名').replace(/[\\/:*?"<>|\s]+/g, '_')}_${new Date(s.createdAt).toISOString().slice(0, 10)}.csv`
  const canShare = !!navigator.canShare?.({ files: [new File(['x'], 'x.csv', { type: 'text/csv' })] })
  sheet(
    `<h2 style="margin:4px 0 14px;font-size:22px">匯出</h2>
     <div class="group">
       <button class="row" id="x-csv"><span class="grow"><span class="title">下載 Excel 檔（CSV）</span><br><span class="meta">${esc(name)}</span></span>${chev}</button>
       ${canShare ? `<button class="row" id="x-share"><span class="grow"><span class="title">分享…</span><br><span class="meta">傳到 LINE、Email、雲端硬碟</span></span>${chev}</button>` : ''}
       <button class="row" id="x-copy"><span class="grow"><span class="title">複製成文字</span><br><span class="meta">直接貼到 LINE 群組</span></span>${chev}</button>
     </div>`,
    (el, close) => {
      el.querySelector('#x-csv').onclick = () => {
        const a = document.createElement('a')
        a.href = URL.createObjectURL(new Blob([csvOf(s)], { type: 'text/csv;charset=utf-8' }))
        a.download = name
        a.click()
        setTimeout(() => URL.revokeObjectURL(a.href), 4000)
        close()
        toast('已下載')
      }
      el.querySelector('#x-share')?.addEventListener('click', async () => {
        try {
          await navigator.share({ files: [new File([csvOf(s)], name, { type: 'text/csv' })], title: '盤點結果', text: textOf(s) })
          close()
        } catch {
          /* 使用者取消分享 */
        }
      })
      el.querySelector('#x-copy').onclick = async () => {
        try {
          await navigator.clipboard.writeText(textOf(s))
          toast('已複製')
        } catch {
          toast('這個瀏覽器不讓複製，請改用下載')
        }
        close()
      }
    },
  )
}

// ───────────────────────── 動作 ─────────────────────────
async function runAnalysis(indices) {
  const s = state.session
  state.cancel = false
  state.progress = { done: 0, total: indices.length, started: Date.now() }
  go('analyzing')
  const timer = setInterval(() => {
    const el = document.getElementById('elapsed')
    if (el) el.textContent = `已經 ${Math.round((Date.now() - state.progress.started) / 1000)} 秒`
  }, 500)
  for (const i of indices) {
    if (state.cancel) break
    const photo = s.photos[i]
    try {
      const r = await analyze(photo, (msg) => {
        const el = document.getElementById('ai-status')
        if (el) el.textContent = msg
      })
      Object.assign(photo, { objects: r.objects, note: r.note, status: 'done', error: '', errorDetail: '' })
      s.model = r.model
    } catch (e) {
      Object.assign(photo, { status: 'error', error: e.message || String(e), errorDetail: e.detail || '' })
    }
    state.progress.done += 1
    await save()
    if (state.view === 'analyzing') render()
  }
  clearInterval(timer)
  go('review', { photoIndex: 0 })
}

async function addFiles(files) {
  const s = state.session
  for (const f of files) {
    try {
      s.photos.push(await prepareImage(f))
    } catch {
      toast('有一張照片讀不進來，換一張試試')
    }
  }
  render()
}

$app.addEventListener('click', async (e) => {
  const t = e.target.closest('[data-go],[data-action],[data-open],[data-remove-photo],[data-photo-index],[data-focus],[data-step],[data-edit],[data-retry]')
  if (!t) {
    // 點空白處取消標示（不捲回頂端）
    if (state.focus && !e.target.closest('.photo-wrap,.item')) {
      state.focus = null
      render()
    }
    return
  }
  const d = t.dataset
  if (d.go) return go(d.go)
  if (d.open) {
    state.session = await db.get(d.open)
    delete state.session.edits // 第一版的舊欄位，不再使用
    return go('review', { photoIndex: 0 })
  }
  if (d.removePhoto !== undefined) {
    state.session.photos.splice(Number(d.removePhoto), 1)
    return render()
  }
  if (d.photoIndex !== undefined) {
    state.photoIndex = Number(d.photoIndex)
    return render()
  }
  if (d.focus !== undefined) {
    // 已經標示的品項，再點照片上的框＝只改那一個框
    if (d.obj && state.focus === d.focus) {
      const [pi, oi] = d.obj.split(':').map(Number)
      return objectSheet(pi, oi)
    }
    const key = state.focus === d.focus ? null : d.focus
    // 點清單時，照片切到有這個品項的那張
    if (key) {
      const g = groupsOf(state.session).find((x) => x.key === key)
      if (g?.refs.length && !g.refs.some((r) => r.pi === state.photoIndex)) state.photoIndex = g.refs[0].pi
    }
    state.focus = key
    render()
    if (key && t.closest('.photo-wrap')) document.querySelector(`[data-item="${CSS.escape(key)}"]`)?.scrollIntoView({ block: 'center', behavior: 'smooth' })
    return
  }
  if (d.step) {
    const s = state.session
    const g = groupsOf(s).find((x) => x.key === d.key)
    if (!g) return
    setCount(s, g, (Number(g.count) || 0) + Number(d.step))
    await save()
    return render()
  }
  if (d.edit) return editSheet(d.edit)
  if (d.retry !== undefined) return runAnalysis([Number(d.retry)])

  switch (d.action) {
    case 'new':
      state.session = { id: uid(), createdAt: Date.now(), place: '', photos: [], counts: {}, manual: [] }
      return go('capture')
    case 'analyze':
      state.session.place = document.getElementById('place')?.value.trim() || ''
      return runAnalysis(state.session.photos.map((_, i) => i))
    case 'cancel':
      state.cancel = true
      return toast('這張辨識完就停')
    case 'add':
      return addSheet()
    case 'export':
      return exportSheet()
    case 'delete-session':
      if (!confirm('確定刪除這次盤點？照片和結果都會刪掉。')) return
      await db.del(state.session.id)
      toast('已刪除')
      return go('home')
    case 'save-key': {
      const v = document.getElementById('apikey').value.trim()
      ls.set(LS.key, v)
      ls.set(LS.model, '')
      if (!v) return toast('已清除 API Key')
      try {
        const list = await fetchModels()
        if (!list.length) return toast('Key 可以用，但找不到能看圖的模型')
        ls.set(LS.model, list[0])
        toast(`可以用了！模型：${list[0]}`)
        return render()
      } catch (err) {
        return toast(err.message)
      }
    }
    case 'toggle-key': {
      const el = document.getElementById('apikey')
      el.type = el.type === 'password' ? 'text' : 'password'
      return
    }
    case 'pick-model':
      try {
        const list = await fetchModels()
        document.getElementById('models').innerHTML = `<div class="group" style="margin-top:10px">${list
          .slice(0, 8)
          .map((m) => `<button class="row" data-action="use-model" data-model="${esc(m)}"><span class="grow">${esc(m)}</span>${m === ls.get(LS.model) ? '✓' : ''}</button>`)
          .join('')}</div><p class="footnote">排越前面越推薦（新、穩定、快）。</p>`
      } catch (err) {
        toast(err.message)
      }
      return
    case 'use-model':
      ls.set(LS.model, d.model)
      toast(`改用 ${d.model}`)
      return render()
    case 'save-catalog':
      ls.set(LS.catalog, document.getElementById('catalog').value.trim())
      return toast('清單已儲存')
    case 'reset-catalog':
      ls.set(LS.catalog, '')
      toast('已恢復預設清單')
      return render()
    case 'clear-all':
      if (!confirm('確定刪除全部盤點紀錄？刪了救不回來。')) return
      await db.clear()
      return toast('已全部刪除')
  }
})

function bindInputs() {
  document.getElementById('cam')?.addEventListener('change', (e) => addFiles([...e.target.files]))
  document.getElementById('pick')?.addEventListener('change', (e) => addFiles([...e.target.files]))
  document.getElementById('place')?.addEventListener('input', (e) => (state.session.place = e.target.value))
  document.querySelectorAll('[data-count]').forEach((input) =>
    input.addEventListener('change', async (e) => {
      const s = state.session
      const g = groupsOf(s).find((x) => x.key === e.target.dataset.count)
      if (g) setCount(s, g, parseInt(e.target.value, 10) || 0)
      await save()
      render()
    }),
  )
}

// 離線也能打開（服務工作程式快取 App 本身；AI 辨識還是要網路）
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {})

render()
