/**
 * 拍照盤點（冷凍材料行）
 * 拍貨架 → Gemini 視覺模型找出每個商品並框起來 → 原圖對照、＋／－ 修正 → 存在手機、匯出 CSV。
 * 沒有後端：API Key 只存在這支手機（localStorage），照片只送到 Google Gemini 分析。
 */

const API = 'https://generativelanguage.googleapis.com/v1beta'
const LS = { key: 'inventory:apiKey', model: 'inventory:model', catalog: 'inventory:catalog', pinned: 'inventory:modelPinned' }
const MAX_SIDE = 1600 // 照片先縮到長邊 1600px 再上傳：夠看清楚，又快
/** 版本：設定頁最下面會顯示，用來確認手機拿到的是新版 */
const VERSION = '1.6（10/2 夜・加快＋雙保險）'

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
銅管接頭（彎頭、三通、直接頭；等徑、異徑，例如 4分×3分）
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
/** AI 有時把清單整行（含括號說明）當品名，例如「銅管接頭（彎頭、三通…）」→ 只留「銅管接頭」 */
const cleanLabel = (label) => {
  const s = String(label ?? '').trim()
  const base = s.replace(/\s*[（(].*$/, '').trim()
  return base && base !== s && catalogNames().includes(base) ? base : s
}

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
/** sessions＝盤點紀錄；samples＝樣品照（第 2 版新增） */
const idb = (() => {
  let p
  const open = () =>
    (p ??= new Promise((resolve, reject) => {
      const req = indexedDB.open('inventory', 2)
      req.onupgradeneeded = () => {
        const d = req.result
        if (!d.objectStoreNames.contains('sessions')) d.createObjectStore('sessions', { keyPath: 'id' })
        if (!d.objectStoreNames.contains('samples')) d.createObjectStore('samples', { keyPath: 'id' })
      }
      req.onblocked = () => toast('App 更新了：請關掉其他開著盤點 App 的分頁，再重新整理')
      req.onsuccess = () => {
        // 別的分頁要升級資料庫時，這邊先關掉，不要卡住它
        req.result.onversionchange = () => req.result.close()
        resolve(req.result)
      }
      req.onerror = () => reject(req.error)
    }))
  const tx = async (store, mode, fn) => {
    const d = await open()
    return new Promise((resolve, reject) => {
      const t = d.transaction(store, mode)
      const r = fn(t.objectStore(store))
      t.oncomplete = () => resolve(r?.result)
      t.onerror = () => reject(t.error)
    })
  }
  const storeOf = (store) => ({
    all: async () => ((await tx(store, 'readonly', (s) => s.getAll())) ?? []).sort((a, b) => b.createdAt - a.createdAt),
    get: (id) => tx(store, 'readonly', (s) => s.get(id)),
    put: (item) => tx(store, 'readwrite', (s) => s.put(item)),
    del: (id) => tx(store, 'readwrite', (s) => s.delete(id)),
    clear: () => tx(store, 'readwrite', (s) => s.clear()),
  })
  return { sessions: storeOf('sessions'), samples: storeOf('samples') }
})()
const db = idb.sessions

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
/** 樣品照縮到長邊 512px：AI 看得出形狀和粗細就好，送多張也不會太慢 */
const SAMPLE_SIDE = 512
async function drawToBlob(bmp, sx, sy, sw, sh, maxSide) {
  const scale = Math.min(1, maxSide / Math.max(sw, sh))
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.round(sw * scale))
  canvas.height = Math.max(1, Math.round(sh * scale))
  canvas.getContext('2d').drawImage(bmp, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height)
  return new Promise((r) => canvas.toBlob(r, 'image/jpeg', 0.85))
}
/** 從盤點照片切下一個框（四周多留一點邊），當樣品照 */
async function cropBox(photo, box) {
  const bmp = await createImageBitmap(photo.blob)
  const [y1, x1, y2, x2] = box.map((n) => n / 1000)
  const padX = (x2 - x1) * 0.08
  const padY = (y2 - y1) * 0.08
  const sx = Math.max(0, (x1 - padX) * bmp.width)
  const sy = Math.max(0, (y1 - padY) * bmp.height)
  const ex = Math.min(bmp.width, (x2 + padX) * bmp.width)
  const ey = Math.min(bmp.height, (y2 + padY) * bmp.height)
  const blob = await drawToBlob(bmp, sx, sy, ex - sx, ey - sy, SAMPLE_SIDE)
  bmp.close?.()
  return blob
}
/** 直接拍一張樣品照（整張縮小） */
async function sampleFromFile(file) {
  const bmp = await createImageBitmap(file, { imageOrientation: 'from-image' }).catch(() => createImageBitmap(file))
  const blob = await drawToBlob(bmp, 0, 0, bmp.width, bmp.height, SAMPLE_SIDE)
  bmp.close?.()
  return blob
}
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
/** timeoutMs：等太久就放棄（之後會自動重試或換模型），不要讓人一直乾等 */
async function call(path, opts = {}, timeoutMs = 0) {
  const key = ls.get(LS.key)
  let res
  const ctrl = new AbortController()
  const timer = timeoutMs ? setTimeout(() => ctrl.abort(), timeoutMs) : 0
  // opts.signal：另一個模型已經先回答了，這邊就取消
  const outer = opts.signal
  const onOuter = () => ctrl.abort()
  outer?.addEventListener('abort', onOuter)
  try {
    if (outer?.aborted) throw new Error('aborted')
    res = await fetch(`${API}/${path}${path.includes('?') ? '&' : '?'}key=${encodeURIComponent(key)}`, { ...opts, signal: ctrl.signal })
  } catch {
    if (outer?.aborted) throw new ApiError('已取消', 499, 'cancelled')
    if (ctrl.signal.aborted) throw new ApiError(`等了 ${Math.round(timeoutMs / 1000)} 秒還沒回，自動重試。`, 408, `timeout ${timeoutMs / 1000}s`)
    throw new ApiError('沒有網路，或連不到 Google。辨識時請不要切到別的 App、不要讓螢幕暗掉。', 0, 'network error')
  } finally {
    clearTimeout(timer)
    outer?.removeEventListener('abort', onOuter)
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
/** 這些狀況值得自動重試：太多人用、伺服器暫時錯誤、額度每分鐘上限、等太久、網路斷一下 */
const retryable = (e) => e instanceof ApiError && [0, 408, 429, 500, 502, 503, 504].includes(e.status)

/**
 * 思考設定：新的 Gemini 預設會先「想」很久才回答（一張照片可以拖到一分鐘）。
 * 找東西、畫框不用想太多 → 調到最少；第二輪比對相似品才讓它想一點點。
 */
function thinkingFor(model, purpose = 'detect') {
  if (/2\.5-flash/.test(model)) return { thinkingBudget: purpose === 'detect' ? 0 : 512 }
  if (/2\.5-pro/.test(model)) return { thinkingBudget: 128 }
  if (/gemini-[3-9]/.test(model)) return { thinkingLevel: purpose === 'detect' && /flash/.test(model) ? 'minimal' : 'low' }
  return null
}
/** 單次等待上限：第一輪找東西、第二輪比對 */
const TIMEOUT = { detect: 40000, refine: 35000 }

/** 自動挑模型：可用、會看圖、名字有 flash；穩定版優先、版本新的優先、lite 排後面 */
function rankModels(models, kind = /flash/) {
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
    .filter((n) => kind.test(n) && !/image|tts|audio|live|embedding|thinking|8b|computer|robotics/.test(n))
    .sort((a, b) => score(b) - score(a))
}
let modelCache = null
/** Pro：看得比較細（相似品分得比較開），但比較慢、免費額度比 Flash 少很多；只在設定裡手動選 */
let proCache = []
async function fetchModels() {
  const data = await call('models?pageSize=200')
  modelCache = rankModels(data.models || [])
  proCache = rankModels(data.models || [], /-pro/)
  return modelCache
}
async function currentModel(force = false) {
  const saved = ls.get(LS.model)
  // 之前忙線時自動換成 Lite 的，下次重新挑（Lite 只當備用）；自己在設定選的就照用
  if (saved && !force && (!/lite/.test(saved) || ls.get(LS.pinned) === saved)) return saved
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

function prompt(sampleCount = 0) {
  return `你是冷凍空調材料行的盤點助手。請找出照片裡每一個「商品」，每一個各給一個框。
規則：
1. 名稱優先用下面「店內品項清單」的寫法；清單沒有就用最具體的中文名稱（例如「乾燥過濾器」，不要只寫「零件」）。
2. 標籤看得到就分開填：brand＝品牌（例如 Danfoss）、model＝型號（例如 DML 083、KP 15）、spec＝尺寸／容量／規格（例如 3分、10.9kg、4L、R404A）。看不出來就空字串，不要猜。
3. 同一種東西有幾個就給幾個框，不要合併成一個；品牌、型號或尺寸不同就算不同的東西。被擋住一半以上的也要算，但 confidence 給低一點。
4. 貨架、標價牌、手、紙箱外的雜物不算商品。
5. box_2d 用 [ymin, xmin, ymax, xmax]，是相對整張照片的 0～1000。
6. 外觀很像、只差尺寸或形狀的同類商品（例如銅管三通、彎頭、接頭），要逐個比較：管子粗細、三個接口是不是一樣粗（等徑）還是有一個比較細（異徑）、長短。不一樣的就分成不同品項，spec 寫出差別（例如「等徑・粗」「等徑・細」「異徑・側口細」）；看得出分數（2分、3分、4分…）就寫分數。不要把不同尺寸全部歸成同一種。
${
  sampleCount
    ? `7. 前面附了 ${sampleCount} 張「店內樣品」照片，每張只拍一個商品，名稱是正確的。【要盤點的照片】裡的商品如果跟某張樣品一樣（形狀、粗細比例、接口大小都一樣），label、brand、model、spec 就照那張樣品填，一字不差；跟每張樣品都不像，才照一般規則寫。
8. box_2d 只標【要盤點的照片】裡的位置；樣品照不要框、不要算數量。
`
    : ''
}店內品項清單：
${catalogLines().join('\n')}`
}

/** 簡化模式不用 responseSchema，改在文字裡說明要的格式 */
const JSON_HINT = `
只回 JSON，不要其他文字，格式：
{"objects":[{"label":"品名","brand":"品牌","model":"型號","spec":"尺寸／規格","box_2d":[ymin,xmin,ymax,xmax],"confidence":0.9}],"note":"需要人工確認的地方"}`

/**
 * 用某個模型送一次辨識。
 * simple＝簡化請求（不帶 responseSchema、不帶思考設定）：有些時候完整設定會讓 Google 一直回 500，簡化後就過了。
 */
async function generate(model, image, simple = false, refs = [], thinking = true, signal = undefined) {
  // 有樣品照：先放樣品（每張前面寫正確名稱），最後才放要盤點的照片，緊接著說明
  const sampleParts = refs.length
    ? [
        { text: '【店內樣品】每張只拍一個商品，名稱是正確的：' },
        ...refs.flatMap((r, i) => [{ text: `樣品 ${i + 1}：${[r.label, r.brand, r.model, r.spec].map((v) => v || '—').join('｜')}` }, { inline_data: { mime_type: 'image/jpeg', data: r.data } }]),
        { text: '【要盤點的照片】' },
      ]
    : []
  const body = {
    contents: [{ role: 'user', parts: [...sampleParts, { inline_data: { mime_type: 'image/jpeg', data: image } }, { text: prompt(refs.length) + (simple ? JSON_HINT : '') }] }],
    generationConfig: simple
      ? { temperature: 0, responseMimeType: 'application/json' }
      : {
          temperature: 0,
          responseMimeType: 'application/json',
          responseSchema: SCHEMA,
          // 思考調到最少：找東西、畫框不用想，速度快很多
          ...(thinking && thinkingFor(model) ? { thinkingConfig: thinkingFor(model) } : {}),
        },
  }
  const data = await call(`models/${model}:generateContent`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body), signal }, TIMEOUT.detect)
  const cand = data?.candidates?.[0]
  const text = cand?.content?.parts?.map((p) => p.text || '').join('') || ''
  if (!text) throw new ApiError(cand?.finishReason === 'SAFETY' ? 'AI 拒絕分析這張照片，請換一張。' : 'AI 沒有回傳結果，請再試一次。', 0, `empty response ${cand?.finishReason || ''}`)
  let parsed
  try {
    // 簡化模式偶爾會包 ```json …```，先剝掉
    parsed = JSON.parse(text.replace(/^```(?:json)?\s*|\s*```$/g, ''))
  } catch {
    throw new ApiError('AI 回傳的格式壞掉了，請再試一次。', 0, `bad json: ${text.slice(0, 120)}`)
  }
  return {
    objects: (parsed.objects || [])
      .filter((o) => Array.isArray(o.box_2d) && o.box_2d.length === 4 && o.label)
      .map((o) => ({
        label: cleanLabel(o.label),
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
 * 分析一張照片：
 * 1. 先用完整設定送；伺服器錯誤（5xx）或設定不被接受（400）→ 同一個模型改用簡化請求再試。
 * 2. 還是忙（503／429）→ 等 2 秒再試一次，不行就換下一個模型（例如 Flash → Flash-Lite），不要讓人一直等。
 * 3. 模型下架（404）→ 換下一個並重新挑預設模型。
 * onStatus 用來在畫面上顯示目前在做什麼。
 */
/** 一次最多送幾張樣品照（越多越慢；新存的優先） */
const MAX_SAMPLES = 30
async function sampleRefs() {
  const list = (await idb.samples.all().catch(() => [])).slice(0, MAX_SAMPLES)
  return Promise.all(list.map(async (s) => ({ label: s.label, brand: s.brand, model: s.model, spec: s.spec, data: await blobToBase64(s.blob) })))
}

/** 主模型等這麼久還沒回，就同時請備用模型做，誰先回就用誰 */
const HEDGE_MS = 15000

async function analyze(photo, onStatus = () => {}, refs = []) {
  const image = await blobToBase64(photo.blob)
  const first = await currentModel()
  const list = modelCache ?? (await fetchModels().catch(() => [first]))
  const models = [first, ...list.filter((m) => m !== first)].slice(0, 4)
  // 備用：最快的 Lite（分組不夠細沒關係，背景的相似品比對會用強的模型再分一次）
  const backup = list.find((m) => m !== first && /lite/.test(m)) ?? list.find((m) => m !== first)
  const hedge = new AbortController()
  const primary = attemptChain(models, image, refs, onStatus, hedge.signal)
  if (!backup) return primary.finally(() => hedge.abort())
  let timer = 0
  let started = false
  const second = new Promise((resolve, reject) => {
    timer = setTimeout(() => {
      started = true
      onStatus(`${first} 比較慢，同時請 ${backup} 一起做，誰先好就用誰`)
      generate(backup, image, false, refs, true, hedge.signal).then(resolve, reject)
    }, HEDGE_MS)
    // 主模型在備用開始前就失敗（已經換過所有模型）：不用再等備用
    primary.catch((e) => {
      if (!started) {
        clearTimeout(timer)
        reject(e)
      }
    })
  })
  try {
    return await Promise.any([primary, second])
  } catch (agg) {
    throw agg?.errors?.find((e) => e?.status !== 499) ?? agg?.errors?.[0] ?? agg
  } finally {
    clearTimeout(timer)
    hedge.abort() // 輸的那一個取消掉，不浪費額度
  }
}

/** 依序試：同一個模型（有思考設定 → 拿掉 → 簡化），不行就換下一個模型 */
async function attemptChain(models, image, refs, onStatus, signal) {
  let lastErr
  for (const [mi, model] of models.entries()) {
    if (signal?.aborted) break
    onStatus(mi === 0 ? `用 ${model} 辨識中` : `改用 ${model} 再試`)
    const plan = [
      { simple: false, thinking: true, wait: 0 },
      // 不接受思考設定（400）或網路斷一下：同一個模型、拿掉思考設定再試
      { simple: false, thinking: false, wait: 0 },
      { simple: true, thinking: false, wait: 1500 },
    ]
    for (const [pi, step] of plan.entries()) {
      if (step.wait) {
        onStatus(`Google 忙線，${step.wait / 1000} 秒後再試（${model}）`)
        await sleep(lastErr?.retryAfter ? Math.min(lastErr.retryAfter * 1000, 15000) : step.wait)
      }
      try {
        // 第二次：上一次是 400（不接受思考設定）才拿掉；網路斷一下就照原本的快速設定再送
        const think = pi === 1 ? lastErr?.status !== 400 : step.thinking
        const r = await generate(model, image, step.simple, refs, think, signal)
        // 這個模型比較順：之後先用它（但不要換成 Lite：Lite 分不出相似品，只當備用）
        if (mi > 0 && !/lite/.test(model)) ls.set(LS.model, model)
        return r
      } catch (e) {
        lastErr = e
        if (state.cancel || signal?.aborted) throw e
        if (!(e instanceof ApiError)) throw e
        if (e.status === 404) {
          if (mi === 0) ls.set(LS.model, '')
          break
        }
        // 等太久：同一個模型再等一次也是慢，直接換下一個（比較快的）模型
        if (e.status === 408) {
          onStatus(`${model} 太慢，換比較快的模型`)
          break
        }
        // 400：先拿掉思考設定、再改簡化請求；都不行就是別的問題（Key、照片）
        if (e.status === 400) {
          if (pi < plan.length - 1) continue
          throw e
        }
        if (!retryable(e)) throw e
        if (e.status === 0) onStatus('網路斷了一下，重新送出')
        // 500 一直出現：換模型比較快；503／429／網路：再試一次
        if (e.status === 500 || pi === plan.length - 1) break
      }
    }
  }
  throw lastErr ?? new ApiError('已取消', 499, 'cancelled')
}

// ───────────────────────── 相似品再比對（第二輪） ─────────────────────────
/**
 * 第一輪是整張照片一起看，同類的小東西（三通、彎頭）常被歸成同一種。
 * 第二輪：把同一類的框用「同一個比例」切成小圖並排給 AI 比（東西越大、小圖越大），一個一個比粗細和接口，再分組。
 */
const REFINE_SCHEMA = {
  type: 'OBJECT',
  properties: {
    groups: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          ids: { type: 'ARRAY', items: { type: 'INTEGER' }, description: '這一組的物件編號' },
          label: { type: 'STRING' },
          brand: { type: 'STRING' },
          model: { type: 'STRING' },
          spec: { type: 'STRING', description: '這一組跟其他組的差別，例如 等徑・粗、異徑・側口細、4分×3分' },
          why: { type: 'STRING', description: '一句話說這一組的特徵' },
        },
        required: ['ids', 'label', 'brand', 'model', 'spec', 'why'],
      },
    },
  },
  required: ['groups'],
}
const MAX_REFINE = 24

/** 框的實際大小（像素） */
function boxPx(photo, o) {
  const [y1, x1, y2, x2] = o.box
  const w = ((x2 - x1) / 1000) * photo.w
  const h = ((y2 - y1) / 1000) * photo.h
  return { w, h, long: Math.max(w, h), short: Math.min(w, h) }
}

/**
 * 哪些框值得再比一次：同一個品名有 2 個以上，而且框的大小差很多（差 15% 以上＝可能是不同尺寸）；
 * 或者有這個品名的樣品照。全部大小都一樣（例如一排同樣的過濾器）就不用多花一次。
 */
function refineCandidates(photo, objects, refs = []) {
  const byLabel = new Map()
  objects.forEach((o, i) => {
    const k = norm(o.label)
    if (!byLabel.has(k)) byLabel.set(k, [])
    byLabel.get(k).push(i)
  })
  const out = []
  for (const [k, idx] of byLabel) {
    if (idx.length < 2) continue
    const sizes = idx.map((i) => boxPx(photo, objects[i]))
    const spread = (f) => Math.max(...sizes.map((s) => s[f])) / Math.max(1, Math.min(...sizes.map((s) => s[f])))
    const hasSample = refs.some((r) => norm(r.label) === k)
    if (hasSample || spread('long') > 1.15 || spread('short') > 1.15) out.push(...idx)
  }
  return out.slice(0, MAX_REFINE)
}

/** 用同一個比例切小圖：最大的那個長邊 384px，其他照比例縮（夠看清楚粗細，圖小送得快） */
async function sameScaleCrops(photo, objects, indices) {
  const bmp = await createImageBitmap(photo.blob)
  const boxes = indices.map((i) => {
    const [y1, x1, y2, x2] = objects[i].box.map((n) => n / 1000)
    const padX = (x2 - x1) * 0.1
    const padY = (y2 - y1) * 0.1
    const sx = Math.max(0, (x1 - padX) * bmp.width)
    const sy = Math.max(0, (y1 - padY) * bmp.height)
    const ex = Math.min(bmp.width, (x2 + padX) * bmp.width)
    const ey = Math.min(bmp.height, (y2 + padY) * bmp.height)
    return { sx, sy, sw: ex - sx, sh: ey - sy }
  })
  const scale = Math.min(2, 384 / Math.max(...boxes.map((b) => Math.max(b.sw, b.sh))))
  const crops = []
  for (const b of boxes) {
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(b.sw * scale))
    canvas.height = Math.max(1, Math.round(b.sh * scale))
    canvas.getContext('2d').drawImage(bmp, b.sx, b.sy, b.sw, b.sh, 0, 0, canvas.width, canvas.height)
    const blob = await new Promise((r) => canvas.toBlob(r, 'image/jpeg', 0.9))
    crops.push({ data: await blobToBase64(blob), w: canvas.width, h: canvas.height })
  }
  bmp.close?.()
  return crops
}

/** 第二輪用最強的非 Lite 模型（Lite 看不出粗細差別） */
async function refineModel() {
  const list = modelCache ?? (await fetchModels().catch(() => []))
  const saved = ls.get(LS.model)
  if (saved && !/lite/.test(saved)) return saved
  return list.find((m) => !/lite/.test(m)) ?? saved ?? list[0]
}

/** 把 objects 裡 indices 這些框重新分組；失敗就保留原本的結果 */
async function refineObjects(photo, objects, indices, refs = [], onStatus = () => {}) {
  if (indices.length < 2) return { changed: 0 }
  const crops = await sameScaleCrops(photo, objects, indices)
  const model = await refineModel()
  onStatus(`用 ${model} 比對相似品（${indices.length} 個）`)
  const sampleParts = refs.length
    ? [
        { text: '【店內樣品】每張只拍一個商品，名稱是正確的：' },
        ...refs.flatMap((r, i) => [{ text: `樣品 ${i + 1}：${[r.label, r.brand, r.model, r.spec].map((v) => v || '—').join('｜')}` }, { inline_data: { mime_type: 'image/jpeg', data: r.data } }]),
      ]
    : []
  const objectParts = [
    { text: '【要比對的物件】每張小圖都是從同一張照片、用同一個比例切下來的：小圖越大，東西越大。' },
    ...crops.flatMap((c, k) => [{ text: `物件 ${k + 1}（小圖 ${c.w}×${c.h} px）` }, { inline_data: { mime_type: 'image/jpeg', data: c.data } }]),
  ]
  const text = `這些物件在照片裡都被認成「${objects[indices[0]].label}」這一類。請一個一個仔細比較，再分組：
1. 形狀：三通、彎頭、直接頭……
2. 主管有多粗：比較小圖的大小和管口的圓有多大（小圖是同一個比例）。
3. 每個接口是不是一樣粗：一樣粗＝等徑；有一個比較細＝異徑。
4. 長短比例。
完全一樣的才放同一組；只要形狀、粗細或接口不一樣，就分成不同組。每個物件編號都要出現在剛好一組裡。
每組的 label 是品名（用店內品項清單的寫法，不要把括號裡的說明寫進去）；spec 寫出這組跟其他組的差別（例如「等徑・粗」「等徑・細」「異徑・側口細」，看得出分數就寫「4分」「4分×3分」）；brand、model 看不出來就空字串；why 用一句話說這組的特徵。
${refs.length ? '如果某一組跟某張樣品一樣，label、brand、model、spec 就照那張樣品填，一字不差。\n' : ''}店內品項清單：
${catalogLines().join('\n')}`
  const send = (withThinking) => {
    const think = withThinking ? thinkingFor(model, 'refine') : null
    const body = {
      contents: [{ role: 'user', parts: [...sampleParts, ...objectParts, { text }] }],
      generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: REFINE_SCHEMA, ...(think ? { thinkingConfig: think } : {}) },
    }
    return call(`models/${model}:generateContent`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }, TIMEOUT.refine)
  }
  let data
  try {
    data = await send(true)
  } catch (e) {
    // 不接受思考設定（400）或網路斷一下：拿掉思考設定再送一次
    if (!(e instanceof ApiError) || ![0, 400].includes(e.status)) throw e
    data = await send(false)
  }
  const raw = data?.candidates?.[0]?.content?.parts?.map((p) => p.text || '').join('') || ''
  const parsed = JSON.parse(raw.replace(/^```(?:json)?\s*|\s*```$/g, ''))
  let changed = 0
  let used = 0
  for (const g of parsed.groups || []) {
    if ((g.ids || []).some((id) => indices[id - 1] !== undefined)) used += 1
    for (const id of g.ids || []) {
      const idx = indices[id - 1]
      if (idx === undefined) continue
      const o = objects[idx]
      const next = { label: cleanLabel(g.label) || o.label, brand: String(g.brand || '').trim(), model: String(g.model || '').trim(), spec: String(g.spec || '').trim() }
      if (FIELDS.some((f) => (o[f] || '') !== next[f])) changed += 1
      Object.assign(o, next, { refined: true })
    }
  }
  return { changed, groups: used, model }
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
      <div class="row"><span>④ 長得很像、只差尺寸的（例如三通），每一種先存一張「樣品照」，AI 就會照著分；或同一種放一起拍</span></div>
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
    ${state.refining && state.refining.session === s.id ? `<div class="hint-card" role="status" style="margin-top:12px"><b>正在比對相似品…</b>（${state.refining.done}/${state.refining.total} 張）好了會自動更新，可以先看結果。</div>` : ''}
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

async function viewSettings() {
  const key = ls.get(LS.key)
  const model = ls.get(LS.model)
  const samples = await idb.samples.all().catch(() => [])
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
    <div class="group"><div class="row"><span class="grow"><span class="title">${esc(model || '自動挑選')}</span><br><span class="meta">自動挑能看圖、最新又快的 Flash；被下架會自動換。相似品分不開時，可以改用 Pro</span></span><button class="btn small secondary" data-action="pick-model">重新挑選</button></div></div>
    <div id="models"></div>
    <p class="section-title">店內品項清單</p>
    <textarea class="field" id="catalog" spellcheck="false">${esc(catalogLines().join('\n'))}</textarea>
    <p class="footnote">一行一種品項，括號裡寫規格或別名。AI 會照這裡的名稱寫，修正時也會跳出來給你選。</p>
    <div class="row-actions" style="margin-top:10px"><button class="btn small" data-action="save-catalog">儲存清單</button><button class="btn small secondary" data-action="reset-catalog">恢復預設</button></div>
    <p class="section-title">樣品照（${samples.length}）</p>
    ${
      samples.length
        ? `<div class="group">${samples
            .map(
              (s, i) =>
                `<div class="row"><img src="${sampleUrl(s)}" alt="" style="width:52px;height:52px;border-radius:10px;object-fit:contain;background:var(--card-2)"><span class="grow"><span class="title">${esc(s.label)}</span><br><span class="meta">${esc(detailOf(s) || '沒有填品牌、型號、尺寸')}${i >= MAX_SAMPLES ? '・超過 30 張，這張不會送' : ''}</span></span><button class="btn small danger" data-action="del-sample" data-id="${esc(s.id)}">刪除</button></div>`,
            )
            .join('')}</div>`
        : ''
    }
    <div class="row-actions" style="margin-top:10px"><label class="btn small secondary">📷 拍一張樣品<input type="file" accept="image/*" capture="environment" id="sample-cam" class="sr-only"></label></div>
    <p class="footnote">長得很像、只差尺寸的商品（例如不同分數的三通），每一種存一張樣品照，名稱和尺寸寫清楚。辨識時會一起送給 AI 比對（最多 30 張，新的優先）。<br>最快的存法：盤點結果裡先點那個框、再點一次 →「儲存，並存成樣品照」。</p>
    <p class="section-title">連線測試</p>
    <div class="stack"><button class="btn small secondary" data-action="diagnose">測試連線</button><div id="diag"></div></div>
    <p class="footnote">辨識一直失敗時按這個，把結果截圖給我看。</p>
    <p class="section-title">資料</p>
    <div class="row-actions"><button class="btn small danger" data-action="clear-all">刪除全部盤點紀錄</button></div>
    <p class="footnote">紀錄（含照片）只存在這支手機的瀏覽器裡；要留底請用「匯出」。</p>
    <p class="footnote" style="margin-top:18px;text-align:center">拍照盤點 版本 ${VERSION}</p>
  </main>`
}

async function render() {
  const html = state.view === 'home' ? await viewHome() : state.view === 'capture' ? viewCapture() : state.view === 'analyzing' ? viewAnalyzing() : state.view === 'review' ? viewReview() : await viewSettings()
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
     <div class="row-actions" style="margin-top:16px"><button class="btn" style="flex:1" id="e-save">儲存</button><button class="btn danger" id="e-del">刪掉</button></div>
     ${g.manual || g.boxes < 2 ? '' : `<button class="btn secondary block" id="e-refine" style="margin-top:10px">再比對一次：把這 ${g.boxes} 個分得更細</button><p class="footnote" style="margin:8px 2px 0">把這一列的框切成小圖並排給 AI，一個一個比粗細和接口（要網路，約 10～30 秒）。</p>`}
     ${g.manual ? '' : '<button class="btn secondary block" id="e-sample" style="margin-top:10px">儲存，並存成樣品照</button><p class="footnote" style="margin:8px 2px 0">用照片上第一個框當樣品。這一列混了不同尺寸的話，請先點照片上那一個框、再點一次，從那裡存。</p>'}`,
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
      el.querySelector('#e-refine')?.addEventListener('click', async (ev) => {
        const btn = ev.currentTarget
        btn.disabled = true
        btn.textContent = '比對中…（約 10～30 秒）'
        try {
          const refs = await sampleRefs()
          // 同一張照片的框一起比（大小比較才有意義）
          const byPhoto = new Map()
          g.refs.forEach((r) => byPhoto.set(r.pi, [...(byPhoto.get(r.pi) || []), r.oi]))
          let groups = 0
          for (const [pi, ois] of byPhoto) {
            const rr = await refineObjects(s.photos[pi], s.photos[pi].objects, ois, refs, (m) => (btn.textContent = m))
            groups = Math.max(groups, rr.groups || 0)
          }
          // 原本手動改過的數量是給「一整列」的，分組後不再適用
          if (s.counts) delete s.counts[g.key]
          state.focus = null
          await save()
          close()
          render()
          toast(groups > 1 ? `分成 ${groups} 種了，請對照照片確認` : 'AI 比對後還是認為是同一種；可以點框個別修改，或存樣品照')
        } catch (e) {
          btn.disabled = false
          btn.textContent = `再比對一次：把這 ${g.boxes} 個分得更細`
          toast(e.message || '比對失敗，請再試一次')
        }
      })
      el.querySelector('#e-sample')?.addEventListener('click', async () => {
        const v = readFields(el)
        if (!v.label) return toast('品名不能空白')
        const first = g.refs[0]
        const box = s.photos[first.pi].objects[first.oi].box
        moveObjects(s, g.refs, v)
        await saveSample(v, await cropBox(s.photos[first.pi], box))
        state.focus = null
        await save()
        close()
        render()
      })
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
     <div class="row-actions" style="margin-top:16px"><button class="btn" style="flex:1" id="o-save">儲存</button><button class="btn danger" id="o-del">這不是商品，刪掉這個框</button></div>
     <button class="btn secondary block" id="o-sample" style="margin-top:10px">儲存，並存成樣品照</button>
     <p class="footnote" style="margin:8px 2px 0">樣品照：AI 以後看到一樣的東西，會照這裡的名稱和尺寸寫。長得很像、只差尺寸的商品，每一種存一張。</p>`,
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
      el.querySelector('#o-sample').onclick = async () => {
        const v = readFields(el)
        if (!v.label) return toast('品名不能空白')
        moveObjects(s, [{ pi, oi }], v)
        await saveSample(v, await cropBox(s.photos[pi], o.box))
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

// ───────────────────────── 樣品照 ─────────────────────────
async function saveSample(fields, blob) {
  await idb.samples.put({ id: uid(), ...fields, blob, createdAt: Date.now() })
  toast('已存成樣品照：下次辨識會拿來比對')
}
const sampleUrls = new Map()
const sampleUrl = (s) => {
  if (!sampleUrls.has(s.id)) sampleUrls.set(s.id, URL.createObjectURL(s.blob))
  return sampleUrls.get(s.id)
}
/** 直接拍一張樣品照，填名稱後存起來 */
async function newSampleSheet(file) {
  let blob
  try {
    blob = await sampleFromFile(file)
  } catch {
    return toast('照片讀不進來，換一張試試')
  }
  const sug = await suggestions()
  const url = URL.createObjectURL(blob)
  sheet(
    `<h2 class="sheet-title">新增樣品照</h2>
     <img src="${url}" alt="樣品照" style="display:block;max-height:180px;margin:0 auto 12px;border-radius:12px">
     ${fieldsHtml({}, sug)}
     <button class="btn block" id="s-save" style="margin-top:16px">存成樣品照</button>`,
    (el, close) => {
      el.querySelector('#s-save').onclick = async () => {
        const v = readFields(el)
        if (!v.label) return toast('請先填品名')
        await saveSample(v, blob)
        URL.revokeObjectURL(url)
        close()
        render()
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
  keepAwake()
  go('analyzing')
  const timer = setInterval(() => {
    const el = document.getElementById('elapsed')
    if (el) el.textContent = `已經 ${Math.round((Date.now() - state.progress.started) / 1000)} 秒`
  }, 500)
  const refs = await sampleRefs()
  const status = (msg) => {
    const el = document.getElementById('ai-status')
    if (el) el.textContent = msg + (refs.length ? `（附 ${refs.length} 張樣品照）` : '')
  }
  // 多張照片一次送兩張（總時間差不多減半；再多會撞到免費額度每分鐘上限）
  const queue = [...indices]
  const worker = async () => {
    while (queue.length && !state.cancel) {
      const photo = s.photos[queue.shift()]
      try {
        const r = await analyze(photo, status, refs)
        Object.assign(photo, { objects: r.objects, note: r.note, status: 'done', error: '', errorDetail: '' })
        s.model = r.model
      } catch (e) {
        Object.assign(photo, { status: 'error', error: e.message || String(e), errorDetail: e.detail || '' })
      }
      state.progress.done += 1
      await save()
      if (state.view === 'analyzing') render()
    }
  }
  await Promise.all([worker(), worker()])
  clearInterval(timer)
  releaseWakeLock()
  go('review', { photoIndex: 0 })
  // 結果先給你看；相似品在背景再比一次，好了自動更新
  if (!state.cancel) backgroundRefine(s, indices, refs)
}

/**
 * 第二輪在背景跑：同一類、但大小不一樣的框，切小圖並排再比一次。
 * 先記住要比的那幾個框（物件本身），等待時你刪了或改了框也不會對錯位置；你改過的框不會被蓋掉。
 */
async function backgroundRefine(s, indices, refs) {
  const jobs = indices
    .map((i) => s.photos[i])
    .filter((p) => p?.status === 'done')
    .map((p) => ({ photo: p, objs: refineCandidates(p, p.objects, refs).map((oi) => p.objects[oi]) }))
    .filter((j) => j.objs.length >= 2)
  if (!jobs.length) return
  state.refining = { session: s.id, total: jobs.length, done: 0 }
  if (state.session === s && state.view === 'review') render()
  let groups = 0
  for (const { photo, objs } of jobs) {
    const free = objs.filter((o) => !o.edited && photo.objects.includes(o))
    try {
      const rr = await refineObjects(photo, free, free.map((_, k) => k), refs)
      groups += rr.groups || 0
      if (rr.model && !/比對/.test(s.model || '')) s.model = rr.model === s.model ? `${s.model}（含相似品比對）` : `${s.model} ＋ ${rr.model} 比對`
    } catch {
      photo.note = [photo.note, '相似品再比對沒有成功，先保留第一輪的結果；可以點品項的 ✎ →「再比對一次」。'].filter(Boolean).join(' ')
    }
    state.refining.done += 1
    await db.put(s)
    if (state.session === s && state.view === 'review') render()
  }
  state.refining = null
  if (state.session === s && state.view === 'review') {
    render()
    toast(groups > 1 ? '相似品比對完成，已經分開不同尺寸' : '相似品比對完成')
  }
}

// 辨識時讓螢幕保持亮著（螢幕暗掉或切到別的 App，手機會把連線斷掉）
let wakeLock = null
async function keepAwake() {
  try {
    wakeLock = await navigator.wakeLock?.request('screen')
  } catch {
    wakeLock = null
  }
}
function releaseWakeLock() {
  wakeLock?.release?.().catch(() => {})
  wakeLock = null
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
        const rows = (arr) => arr.map((m) => `<button class="row" data-action="use-model" data-model="${esc(m)}"><span class="grow">${esc(m)}</span>${m === ls.get(LS.model) ? '✓' : ''}</button>`).join('')
        document.getElementById('models').innerHTML = `<p class="section-title">快速（Flash）</p><div class="group">${rows(list.slice(0, 6))}</div><p class="footnote">排越前面越推薦（新、穩定、快）。平常用這個。</p>${
          proCache.length
            ? `<p class="section-title">精細（Pro）</p><div class="group">${rows(proCache.slice(0, 3))}</div><p class="footnote">看得比較細，相似品比較分得開；但每張要等比較久，免費額度也比 Flash 少很多。Flash 分不出來、又不想存樣品照時再試。</p>`
            : ''
        }`
      } catch (err) {
        toast(err.message)
      }
      return
    case 'diagnose': {
      const out = document.getElementById('diag')
      const lines = []
      const show = () => (out.innerHTML = `<div class="group">${lines.map((l) => `<div class="row"><span class="grow" style="font-size:14px;word-break:break-all">${l}</span></div>`).join('')}</div>`)
      lines.push('版本 ' + esc(VERSION))
      lines.push('API Key：' + (ls.get(LS.key) ? '有（' + esc(ls.get(LS.key).slice(0, 6)) + '…）' : '<b>沒有</b>'))
      show()
      try {
        const list = await fetchModels()
        lines.push('✅ 模型清單：找到 ' + list.length + ' 個可用（' + esc(list.slice(0, 3).join('、')) + '）')
        show()
        const model = await currentModel()
        const t0 = Date.now()
        const data = await call('models/' + model + ':generateContent', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: '只回答 OK' }] }] }) })
        const reply = data?.candidates?.[0]?.content?.parts?.map((p) => p.text || '').join('') || '（沒有回覆）'
        lines.push('✅ ' + esc(model) + ' 回覆「' + esc(reply.trim().slice(0, 20)) + '」，花 ' + ((Date.now() - t0) / 1000).toFixed(1) + ' 秒')
      } catch (err) {
        lines.push('❌ ' + esc(err.message) + (err.detail ? '<br><span class="muted">' + esc(err.detail) + '</span>' : ''))
      }
      return show()
    }
    case 'use-model':
      ls.set(LS.model, d.model)
      ls.set(LS.pinned, d.model) // 自己選的：之後照用，不會被自動換掉
      toast(`改用 ${d.model}`)
      return render()
    case 'save-catalog':
      ls.set(LS.catalog, document.getElementById('catalog').value.trim())
      return toast('清單已儲存')
    case 'reset-catalog':
      ls.set(LS.catalog, '')
      toast('已恢復預設清單')
      return render()
    case 'del-sample':
      if (!confirm('刪掉這張樣品照？')) return
      await idb.samples.del(d.id)
      toast('已刪掉樣品照')
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
  document.getElementById('sample-cam')?.addEventListener('change', (e) => e.target.files[0] && newSampleSheet(e.target.files[0]))
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
