/**
 * 拍照盤點（冷凍材料行）
 * 拍貨架 → Gemini 視覺模型找出每個商品並框起來 → 原圖對照、＋／－ 修正 → 存在手機；總表一次匯出 Excel、同步 Google 試算表。
 * 品項庫：按「完成」時自動長出來（料號、各位置數量、帳面數、安全庫存）；查型號：拍標籤或打型號 → 解讀、店裡有沒有、替代品。
 * 沒有後端：API Key 只存在這支手機（localStorage），照片只送到 Google Gemini 分析。
 */

import { makeXlsx } from './xlsx.js'
import { decode, normalizeModel, looseKey, canon, modelKey, linksFor } from './rules.js'

const API = 'https://generativelanguage.googleapis.com/v1beta'
const LS = {
  key: 'inventory:apiKey',
  model: 'inventory:model',
  catalog: 'inventory:catalog',
  pinned: 'inventory:modelPinned',
  sheet: 'inventory:sheetUrl',
  autoSync: 'inventory:autoSync',
  locations: 'inventory:locations',
  // 多台裝置同步
  syncKey: 'inventory:syncKey',
  device: 'inventory:deviceId',
  deleted: 'inventory:deleted',
  settingsAt: 'inventory:settingsAt',
  settingsSyncT: 'inventory:settingsSyncT',
  pulled: 'inventory:syncPulled',
  lastSync: 'inventory:lastSync',
  memberName: 'inventory:memberName',
  memberRole: 'inventory:memberRole',
  memberId: 'inventory:memberId',
  roster: 'inventory:roster',
  counterId: 'inventory:counterId',
}
/** 檢視者不能用的動作 */
const EDIT_ACTIONS = new Set(['new', 'analyze', 'add', 'delete-session', 'finish', 'save-catalog', 'reset-catalog', 'add-box', 'del-sample', 'clear-all', 'review-doubts', 'item-add', 'import', 'item-edit', 'item-confirm', 'item-merge', 'item-delete', 'move-in', 'move-out', 'book-set', 'equiv-add', 'read-add', 'loc-add', 'safety-pick'])
/** 權限（跟 Google 雲端硬碟的「共用」一樣） */
const ROLE_LABEL = { owner: '擁有者', manager: '管理員', editor: '編輯者', viewer: '檢視者' }
const ROLE_DESC = { owner: '全部都可以；不能被移除', manager: '可以盤點、修改，也可以邀請、移除人', editor: '可以盤點、修改', viewer: '只能看（可以下載 Excel）' }
const myRole = () => (ls.get(LS.syncKey) ? ls.get(LS.memberRole) || 'editor' : 'owner')
const canEdit = () => myRole() !== 'viewer'
const canManage = () => ['owner', 'manager'].includes(myRole())
/** 共用名單（只有名字）：顯示「誰盤的」用；名字改了，以前的紀錄也顯示新名字 */
const roster = () => {
  try {
    return JSON.parse(ls.get(LS.roster, '[]')) || []
  } catch {
    return []
  }
}
const personName = (id, fallback = '') => roster().find((p) => p.id === id)?.name || fallback
/** 這台的盤點人：預設是這台的使用者；擁有者／管理員可以在設定改（例如店裡共用的平板） */
const currentCounter = () => {
  const id = ls.get(LS.counterId) || ls.get(LS.memberId)
  const name = personName(id, id === ls.get(LS.memberId) ? ls.get(LS.memberName) : '')
  return id && name ? { id, name } : null
}
const byName = (s) => (s.byId ? personName(s.byId, s.by) : s.by) || ''
const MAX_SIDE = 1600 // 照片先縮到長邊 1600px 再上傳：夠看清楚，又快
/** 版本：設定頁最下面會顯示，用來確認手機拿到的是新版 */
const VERSION = '3.5（10/5・試算表：總覽、庫存、儲位庫存、盤差報告）'

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
/** sessions＝盤點紀錄；samples＝樣品照（第 2 版新增）；items＝品項庫（第 3 版新增） */
const idb = (() => {
  let p
  const open = () =>
    (p ??= new Promise((resolve, reject) => {
      const req = indexedDB.open('inventory', 3)
      req.onupgradeneeded = () => {
        const d = req.result
        if (!d.objectStoreNames.contains('sessions')) d.createObjectStore('sessions', { keyPath: 'id' })
        if (!d.objectStoreNames.contains('samples')) d.createObjectStore('samples', { keyPath: 'id' })
        if (!d.objectStoreNames.contains('items')) d.createObjectStore('items', { keyPath: 'id' })
      }
      req.onblocked = () => toast('App 更新了：請關掉其他開著盤點 App 的分頁，再重新整理')
      req.onsuccess = () => {
        // 別的分頁要升級資料庫時，這邊先關掉，不要卡住它；連線斷了，下次用的時候重新打開
        req.result.onversionchange = () => {
          req.result.close()
          p = undefined
        }
        req.result.onclose = () => (p = undefined)
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
    /** 一般存檔：記下修改時間（多台同步靠它判斷哪一份比較新），稍後自動同步 */
    put: (item) => {
      item.updatedAt = Date.now()
      scheduleSync()
      return tx(store, 'readwrite', (s) => s.put(item))
    },
    /** 原樣存（同步下載的、只改本機狀態的）：不改修改時間、不觸發同步 */
    putRaw: (item) => tx(store, 'readwrite', (s) => s.put(item)),
    /** 上傳成功：如果這段時間沒再改過，標記「已同步」（同一個交易裡讀和寫，不會蓋掉剛改的） */
    markSynced: (id, t) =>
      tx(store, 'readwrite', (s) => {
        const req = s.get(id)
        req.onsuccess = () => {
          const v = req.result
          if (v && (v.updatedAt || v.createdAt) === t) s.put({ ...v, _syncT: t })
        }
        return req
      }),
    del: (id) => tx(store, 'readwrite', (s) => s.delete(id)),
    clear: () => tx(store, 'readwrite', (s) => s.clear()),
  })
  return { sessions: storeOf('sessions'), samples: storeOf('samples'), items: storeOf('items') }
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
          odd: { type: 'BOOLEAN', description: '跟同一種的其他個比，大小、粗細、形狀、顏色或標籤明顯不一樣就 true' },
          odd_reason: { type: 'STRING', description: 'odd 是 true 時，10 個字以內寫哪裡不同（例如「比較粗」「側口比較細」）' },
          item_no: { type: 'STRING', description: '就是「品項庫」裡的某一項時填料號（例如 P0012）；不確定就空字串' },
        },
        required: ['label', 'brand', 'model', 'spec', 'box_2d', 'confidence'],
      },
    },
    note: { type: 'STRING', description: '看不清楚、被擋住、需要人工確認的地方；沒有就空字串' },
    location: { type: 'STRING', description: '照片裡看得到儲位標籤（例如 A-01）就填標籤上的代號；沒有就空字串' },
  },
  required: ['objects', 'note'],
}

/** 品項庫裡確認過的（最近用到的優先，最多 150 個）：給 AI 照抄名稱、填料號 */
const promptItems = () =>
  (itemsCache || [])
    .filter((it) => it.status === 'ok')
    .sort((a, b) => (lastCounted(b) || b.updatedAt) - (lastCounted(a) || a.updatedAt))
    .slice(0, 150)
const itemsPrompt = () => {
  const list = promptItems()
  return list.length
    ? `\n品項庫（料號｜品名｜品牌｜型號｜尺寸／規格；照片裡的東西就是其中一項時，item_no 填料號，其他欄位照抄；不確定就 item_no 空字串，照一般規則寫）：\n${list.map((it) => [it.no, it.label, it.brand || '—', it.model || '—', it.spec || '—'].join('｜')).join('\n')}`
    : ''
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
7. 同一種有好幾個時，一個一個看：如果某一個的大小、粗細、接口、顏色或標籤跟其他個不一樣（你不確定是不是同一種），那一個的 odd 設 true，odd_reason 用 10 個字以內寫哪裡不同。一樣的就不用寫 odd。
8. 同一個東西只給一個框，不要重複框。
${
  sampleCount
    ? `9. 前面附了 ${sampleCount} 張「店內樣品」照片，每張只拍一個商品，名稱是正確的。【要盤點的照片】裡的商品如果跟某張樣品一樣（形狀、粗細比例、接口大小都一樣），label、brand、model、spec 就照那張樣品填，一字不差；跟每張樣品都不像，才照一般規則寫。
10. box_2d 只標【要盤點的照片】裡的位置；樣品照不要框、不要算數量。
`
    : ''
}・照片裡看得到儲位標籤（白底黑字的代號，例如 A-01、B-12）時，location 填標籤上的代號；標籤本身不是商品，不要框。
店內品項清單：
${catalogLines().join('\n')}${itemsPrompt()}`
}

/** 簡化模式不用 responseSchema，改在文字裡說明要的格式 */
const JSON_HINT = `
只回 JSON，不要其他文字，格式：
{"objects":[{"label":"品名","brand":"品牌","model":"型號","spec":"尺寸／規格","box_2d":[ymin,xmin,ymax,xmax],"confidence":0.9,"odd":false,"odd_reason":"","item_no":""}],"note":"需要人工確認的地方","location":""}`

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
  // 名稱統一：AI 說是品項庫的某一項，或寫法跟品項庫的某一項一樣（例如「3分 三通」＝「三通 3分」）→ 照品項庫寫；
  // 同一張照片裡寫法不同但其實一樣的，照第一個寫。同一種東西每次寫法都一樣，清單和總表才不會分成兩種。
  const byNo = new Map(promptItems().map((it) => [it.no.toLowerCase(), it]))
  const firstOf = new Map()
  const objects = (parsed.objects || [])
    .filter((o) => Array.isArray(o.box_2d) && o.box_2d.length === 4 && o.label)
    .map((o) => {
      const raw = { label: cleanLabel(o.label), brand: String(o.brand || '').trim(), model: normalizeModel(o.model), spec: String(o.spec || '').trim() }
      const it = byNo.get(String(o.item_no || '').trim().toLowerCase()) || (itemsCache && findItem(itemsCache, raw))
      const k = looseKey(raw)
      if (!firstOf.has(k)) firstOf.set(k, raw)
      return {
        ...(it ? { label: it.label, brand: it.brand, model: it.model, spec: it.spec } : firstOf.get(k)),
        box: o.box_2d.map((n) => Math.min(1000, Math.max(0, Number(n) || 0))),
        confidence: Math.min(1, Math.max(0, Number(o.confidence) || 0)),
        ...(o.odd ? { odd: true, oddReason: String(o.odd_reason || '').trim().slice(0, 20) } : {}),
      }
    })
  const { kept, merged } = dedupe(objects)
  const note = [String(parsed.note || ''), merged ? `同一個東西被框了兩次的，已經合併 ${merged} 個。` : ''].filter(Boolean).join(' ')
  return { objects: kept, note, model, location: String(parsed.location || '').trim() }
}

/** 兩個框重疊 7 成以上、而且是同一種：AI 把同一個東西框了兩次 → 留比較有把握的那個（數量才不會多算） */
function iou(a, b) {
  const [ay1, ax1, ay2, ax2] = a
  const [by1, bx1, by2, bx2] = b
  const iw = Math.max(0, Math.min(ax2, bx2) - Math.max(ax1, bx1))
  const ih = Math.max(0, Math.min(ay2, by2) - Math.max(ay1, by1))
  const inter = iw * ih
  const union = (ax2 - ax1) * (ay2 - ay1) + (bx2 - bx1) * (by2 - by1) - inter
  return union > 0 ? inter / union : 0
}
function dedupe(objects) {
  const sorted = [...objects].sort((a, b) => b.confidence - a.confidence)
  const kept = []
  for (const o of sorted) if (!kept.some((k) => norm(k.label) === norm(o.label) && iou(k.box, o.box) > 0.7)) kept.push(o)
  // 保持原本的順序
  const keep = new Set(kept)
  return { kept: objects.filter((o) => keep.has(o)), merged: objects.length - kept.length }
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
      let t0 = Date.now()
      if (step.wait) {
        onStatus(`Google 忙線，${step.wait / 1000} 秒後再試（${model}）`)
        await sleep(lastErr?.retryAfter ? Math.min(lastErr.retryAfter * 1000, 15000) : step.wait)
      }
      try {
        // 第二次：上一次是 400（不接受思考設定）才拿掉；網路斷一下就照原本的快速設定再送
        const think = pi === 1 ? lastErr?.status !== 400 : step.thinking
        t0 = Date.now()
        const r = await generate(model, image, step.simple, refs, think, signal)
        onStatus(`${model} 完成（${Math.round((Date.now() - t0) / 1000)} 秒，找到 ${r.objects.length} 個）`)
        // 這個模型比較順：之後先用它（但不要換成 Lite：Lite 分不出相似品，只當備用）
        if (mi > 0 && !/lite/.test(model)) ls.set(LS.model, model)
        return r
      } catch (e) {
        lastErr = e
        if (state.cancel || signal?.aborted) throw e
        onStatus(`${model} 沒成功：${e.detail || e.message}（${Math.round((Date.now() - t0) / 1000)} 秒）`)
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
  // 物件很多時小圖再小一點（送得快）；還是同一個比例
  const side = indices.length > 12 ? 256 : 384
  const scale = Math.min(2, side / Math.max(...boxes.map((b) => Math.max(b.sw, b.sh))))
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
  // 物件越多，給 AI 的時間越多（背景在跑，不會卡住畫面）
  const timeout = Math.min(60000, TIMEOUT.refine + indices.length * 1000)
  const send = (think) => {
    const body = {
      contents: [{ role: 'user', parts: [...sampleParts, ...objectParts, { text }] }],
      generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: REFINE_SCHEMA, ...(think ? { thinkingConfig: think } : {}) },
    }
    return call(`models/${model}:generateContent`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }, timeout)
  }
  const t0 = Date.now()
  let data
  try {
    data = await send(thinkingFor(model, 'refine'))
  } catch (e) {
    if (!(e instanceof ApiError)) throw e
    // 等太久或網路斷：思考再調少一點；不接受思考設定（400）：拿掉再送；其他錯誤不重試
    const retry = [0, 408].includes(e.status) && /gemini-[3-9]/.test(model) ? { thinkingLevel: 'minimal' } : [0, 400, 408].includes(e.status) ? null : undefined
    if (retry === undefined) throw e
    onStatus(`比對第一次沒成功（${e.detail || e.status}），再試一次`)
    data = await send(retry)
  }
  const raw = data?.candidates?.[0]?.content?.parts?.map((p) => p.text || '').join('') || ''
  let parsed
  try {
    parsed = JSON.parse(raw.replace(/^```(?:json)?\s*|\s*```$/g, ''))
  } catch {
    throw new ApiError('比對結果格式壞掉', 0, `bad json: ${raw.slice(0, 80) || data?.candidates?.[0]?.finishReason || 'empty'}`)
  }
  onStatus(`${model} 比對完成（${Math.round((Date.now() - t0) / 1000)} 秒）`)
  let changed = 0
  let used = 0
  for (const g of parsed.groups || []) {
    if ((g.ids || []).some((id) => indices[id - 1] !== undefined)) used += 1
    for (const id of g.ids || []) {
      const idx = indices[id - 1]
      if (idx === undefined) continue
      const o = objects[idx]
      // 比對要等一陣子：這段時間你手動改過的框，以你改的為準，不要蓋掉
      if (o.edited) continue
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
  // 編號和顏色固定：第一次出現的順序記在 session.order／session.colorOf；新出現的種類排在最後、拿新的顏色。
  // （以前照名稱排序：新增一種「銅管接頭」會插到「銅管接頭・三通」前面，原本的 3 號變 4 號、顏色也換掉，很容易看錯）
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  const order = (session.order ??= [])
  const colorOf = (session.colorOf ??= {})
  const fresh = [...map.values()].filter((g) => !order.includes(g.key)).sort((a, b) => cmp(a.label, b.label) || cmp(a.spec, b.spec) || cmp(a.brand, b.brand) || cmp(a.model, b.model))
  for (const g of fresh) order.push(g.key)
  const groups = [...map.values()].map((g) => ({ ...g, conf: g.boxes ? g.conf / g.boxes : 1 })).sort((a, b) => order.indexOf(a.key) - order.indexOf(b.key))
  for (const m of session.manual || []) groups.push({ key: `manual:${m.id}`, label: m.label, brand: m.brand || '', model: m.model || '', spec: m.spec || '', boxes: 0, conf: 1, refs: [], manual: m })
  for (const g of groups) {
    if (colorOf[g.key] === undefined) {
      const used = new Set(groups.map((x) => colorOf[x.key]).filter((c) => c !== undefined))
      let c = 0
      while (used.has(c) && c < COLORS.length) c++
      colorOf[g.key] = c % COLORS.length
    }
  }
  return groups.map((g) => ({ ...g, color: COLORS[colorOf[g.key]], count: g.manual ? g.manual.count : (session.counts?.[g.key] ?? g.boxes) }))
}

/**
 * 「要確認」的框：自動挑出可能不一樣的，你只要看這幾個（不用一個一個看）。
 * 理由（照順序判斷，一個框只列一個理由）：
 * ① AI 自己說跟同一種的其他個不一樣（odd）
 * ② 大小跟同一種的其他個差很多（同一張照片裡比；3 個以上差 30%、只有 2 個差 35%）
 * ③ AI 沒把握（信心 < 0.6）
 * 你確認過（按「一樣」或改過種類）就不再列出。
 */
function doubtsOf(session) {
  const out = []
  session.photos.forEach((photo, pi) => {
    if (!photo.w || !photo.h) return
    const byKey = new Map()
    photo.objects.forEach((o) => {
      const k = keyOf(o)
      if (!byKey.has(k)) byKey.set(k, [])
      byKey.get(k).push(o)
    })
    for (const objs of byKey.values()) {
      const sizes = objs.map((o) => boxPx(photo, o).long)
      const med = [...sizes].sort((a, b) => a - b)[Math.floor(sizes.length / 2)] || 1
      const pairGap = objs.length === 2 ? Math.max(...sizes) / Math.max(1, Math.min(...sizes)) : 1
      objs.forEach((o, i) => {
        if (o.checked || o.edited) return
        const diff = sizes[i] / med - 1
        let reason = ''
        if (o.odd) reason = o.oddReason ? `AI 說這一個${o.oddReason}` : 'AI 覺得跟同一種的其他個不太一樣'
        else if (objs.length >= 3 && Math.abs(diff) > 0.3) reason = `比同一種的其他個${diff > 0 ? '大' : '小'}約 ${Math.round(Math.abs(diff) * 100)}%`
        else if (pairGap > 1.35) reason = '這一種只有 2 個，但大小差很多'
        else if (o.confidence < 0.6) reason = 'AI 不太確定這是什麼'
        if (reason) out.push({ pi, o, reason })
      })
    }
  })
  return out
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

/**
 * 首頁每一筆的標題：盤了什麼、幾件（重點），不是位置。
 * 例：「銅管接頭・三通 ×23」「銅管接頭 4 種，共 8 件」「乾燥過濾器、視液鏡等 5 種，共 30 件」
 */
function summaryOf(session) {
  const groups = groupsOf(session).filter((g) => Number(g.count) > 0)
  if (!groups.length) return '沒有找到商品'
  const total = totalQty(session)
  if (groups.length === 1) {
    const g = groups[0]
    return `${g.label}${g.spec ? `・${g.spec}` : ''} ×${total}`
  }
  const labels = [...new Set(groups.map((g) => g.label))]
  if (labels.length === 1) return `${labels[0]} ${groups.length} 種，共 ${total} 件`
  return `${labels.slice(0, 2).join('、')}${labels.length > 2 ? ' 等' : ''} ${groups.length} 種，共 ${total} 件`
}

// ───────────────────────── 儲位：貨架每一格一個代號（A-01） ─────────────────────────
const locations = () => {
  try {
    return JSON.parse(ls.get(LS.locations, '[]')) || []
  } catch {
    return []
  }
}
/** 儲位、品項清單改了：記下時間，稍後同步到別台 */
const touchSettings = () => {
  ls.set(LS.settingsAt, String(Date.now()))
  scheduleSync()
}
const saveLocations = (list) => {
  ls.set(LS.locations, JSON.stringify(list))
  touchSettings()
}
const findLocation = (code) => locations().find((l) => canon(l.code) === canon(code))
/** 顯示用：「A-01（冷凍油那排第 1 層）」 */
const placeLabel = (place) => {
  const loc = findLocation(place)
  return loc ? (loc.name ? `${loc.code}（${loc.name}）` : loc.code) : place
}
/** 下一個代號：A-03 → A-04 */
const nextCode = (list) => {
  const last = list.at(-1)?.code || ''
  const m = /^(.*?)(\d+)$/.exec(last)
  return m ? `${m[1]}${String(Number(m[2]) + 1).padStart(m[2].length, '0')}` : 'A-01'
}

// ───────────────────────── 品項庫：盤點時自動長出來 ─────────────────────────
/**
 * 每一種商品一筆，料號 P0001 起。按「完成」時，這次盤點的每一種都會對到品項庫：
 * 對得到 → 記下這個位置的數量；對不到 → 自動建立「新的」品項（之後在「品項」確認名稱、合併重複的）。
 * 實盤＝每個位置「最近一次」數到的數量加起來；帳面數＝應該要有幾個（進貨加、賣出減，或拿實盤當起點）。
 */
let itemsCache = null
async function itemsAll(force = false) {
  if (!itemsCache || force) itemsCache = await idb.items.all().catch(() => [])
  return itemsCache
}
async function putItem(it) {
  it.updatedAt = Date.now()
  await idb.items.put(it)
  if (itemsCache && !itemsCache.includes(it)) itemsCache.unshift(it)
}
async function delItem(it) {
  await idb.items.del(it.id)
  tombstone(`item:${it.id}`)
  if (itemsCache) itemsCache = itemsCache.filter((x) => x !== it)
}
const nextNo = (items) => `P${String(items.reduce((m, it) => Math.max(m, parseInt(String(it.no).slice(1), 10) || 0), 0) + 1).padStart(4, '0')}`
function newItem(items, f, extra = {}) {
  const now = Date.now()
  // 型號看得懂（例如 DML 083S）：沒寫的品牌、規格自動補上（原本的寫法也記成別名）
  const d = decode(f.model).find((x) => x.model)
  const v = { label: f.label, brand: f.brand || d?.brand || '', model: f.model || '', spec: f.spec || d?.spec || '' }
  return { id: uid(), no: nextNo(items), createdAt: now, updatedAt: now, ...v, aliases: [...new Set([looseKey(f), looseKey(v)])], status: 'new', stock: {}, book: null, safety: null, moves: [], equiv: [], ...extra }
}
/** 找品項：先比寬鬆的名稱（含以前的寫法），再比型號＋品名 */
function findItem(items, f) {
  const k = looseKey(f)
  const hit = items.find((it) => it.aliases?.includes(k))
  if (hit) return hit
  const mk = modelKey(f.model)
  return (mk.length >= 3 && items.find((it) => modelKey(it.model) === mk && canon(it.label) === canon(f.label))) || null
}
/**
 * 各位置的數量：{ count, at（盤點時間）, sid（哪次盤點）, place, v（寫入時間）, removed（這一格已經沒有了） }
 * 拿掉一格不直接刪，而是標 removed：多人同步時，別台舊的資料才不會讓它又冒出來。
 */
const liveStock = (it) => Object.entries(it.stock || {}).filter(([, st]) => !st.removed)
/** 同一格兩個版本：盤點時間比較新的贏；一樣就看寫入時間 */
function newerEntry(a, b) {
  if (!a) return b
  if (!b) return a
  return ((a.at || 0) - (b.at || 0) || (a.v || 0) - (b.v || 0)) >= 0 ? a : b
}
const removedEntry = (st, at = st.at) => ({ ...st, count: 0, removed: true, at: Math.max(st.at || 0, at), v: Date.now() })
const onHand = (it) => liveStock(it).reduce((n, [, st]) => n + (Number(st.count) || 0), 0)
const lastCounted = (it) => Math.max(0, ...liveStock(it).map(([, st]) => st.at || 0))
/** 差異＝實盤－帳面（沒有帳面數或還沒盤過就不算） */
const diffOf = (it) => (it.book == null || !liveStock(it).length ? null : onHand(it) - it.book)
/** 帳面數從進出紀錄算：最後一次「設定」＋之後的進貨－賣出（多人同時記進貨也不會少算） */
function bookFromMoves(moves) {
  const sorted = [...(moves || [])].sort((a, b) => a.at - b.at)
  let book = null
  for (const m of sorted) {
    if (m.kind === 'set') book = m.qty
    else if (book != null) book = Math.max(0, book + (m.kind === 'in' ? m.qty : -m.qty))
  }
  return book
}
/** 現在大概有幾個：有帳面數用帳面數（進貨／賣出會改它），沒有就用實盤 */
const expected = (it) => (it.book != null ? it.book : onHand(it))
const needsOrder = (it) => it.safety != null && expected(it) <= it.safety
const placeKeyOf = (s) => (s.place ? `p:${canon(s.place)}` : `s:${s.id}`)
const stockPlace = (st) => (st.place ? placeLabel(st.place) : `沒填位置・${fmtTime(st.at)}`)
const itemTitle = (it) => `${it.label}${it.spec ? `・${it.spec}` : ''}`
const itemUrls = new Map()
const itemUrl = (it) => {
  if (!it.photo) return ''
  const hit = itemUrls.get(it.id)
  if (hit?.blob === it.photo) return hit.url
  if (hit) URL.revokeObjectURL(hit.url)
  const url = URL.createObjectURL(it.photo)
  itemUrls.set(it.id, { blob: it.photo, url })
  return url
}
const itemThumb = (it, size = 44) =>
  it.photo ? `<img class="thumb" src="${itemUrl(it)}" alt="" style="width:${size}px;height:${size}px">` : `<span class="thumb ph" style="width:${size}px;height:${size}px" aria-hidden="true">${esc(it.label.slice(0, 1))}</span>`

/**
 * 按「完成」：這次盤點記進品項庫。可以重複按（改完再按一次會更新，不會重複算）。
 * 同一個位置，以「比較新的那次盤點」為準。
 */
async function linkSession(s) {
  const items = await itemsAll()
  const pk = placeKeyOf(s)
  const sums = new Map()
  let created = 0
  s.itemOf = {}
  for (const g of groupsOf(s)) {
    const count = Number(g.count) || 0
    let it = findItem(items, g)
    if (!it) {
      if (!count) continue
      it = newItem(items, g)
      const r = g.refs[0]
      if (r) it.photo = await cropBox(s.photos[r.pi], s.photos[r.pi].objects[r.oi].box).catch(() => undefined)
      items.unshift(it)
      created++
    } else {
      const k = looseKey(g)
      if (!it.aliases.includes(k)) it.aliases.push(k)
    }
    s.itemOf[g.key] = it.id
    sums.set(it, (sums.get(it) || 0) + count)
  }
  // 同一格以這次為準：這次沒數到的（以前記在這一格、或這次盤點改成別的）→ 從這一格拿掉
  for (const it of items) {
    const st = it.stock?.[pk]
    if (st && !st.removed && !sums.has(it) && (st.sid === s.id || st.at <= s.createdAt)) {
      it.stock[pk] = removedEntry(st, s.createdAt)
      await putItem(it)
    }
  }
  for (const [it, count] of sums) {
    const st = it.stock[pk]
    if (!st || st.sid === s.id || st.at <= s.createdAt) it.stock[pk] = { count, at: s.createdAt, sid: s.id, place: s.place || '', v: Date.now() }
    await putItem(it)
  }
  s.linkedAt = Date.now()
  await db.put(s)
  return { created, linked: sums.size }
}
/** 刪掉一次盤點：它記在品項庫的數量也拿掉 */
async function unlinkSession(id) {
  for (const it of await itemsAll()) {
    const keys = liveStock(it)
      .filter(([, st]) => st.sid === id)
      .map(([k]) => k)
    if (!keys.length) continue
    keys.forEach((k) => (it.stock[k] = removedEntry(it.stock[k])))
    await putItem(it)
  }
}
/** 合併重複的品項：from 併進 to（以前的寫法、各位置數量、進出紀錄都帶過去） */
async function mergeItems(from, to) {
  to.aliases = [...new Set([...(to.aliases || []), ...(from.aliases || [])])]
  for (const [k, st] of Object.entries(from.stock || {})) {
    const cur = to.stock[k]
    if (cur && !cur.removed && !st.removed && cur.sid === st.sid) to.stock[k] = { ...cur, count: cur.count + st.count, v: Date.now() }
    else to.stock[k] = newerEntry(cur, st)
  }
  if (from.book != null) {
    to.book = (to.book ?? 0) + from.book
    to.moves = [...(to.moves || []), { at: Date.now(), kind: 'set', qty: to.book }] // 合併後的帳面數當新起點
  }
  if (to.safety == null) to.safety = from.safety
  to.moves = [...(to.moves || []), ...(from.moves || [])].sort((a, b) => a.at - b.at)
  to.equiv = [...new Set([...(to.equiv || []), ...(from.equiv || [])])].filter((id) => id !== to.id && id !== from.id)
  if (!to.photo && from.photo) to.photo = from.photo
  to.status = 'ok'
  for (const it of await itemsAll()) {
    if (it.equiv?.includes(from.id)) {
      it.equiv = [...new Set(it.equiv.map((id) => (id === from.id ? to.id : id)))].filter((id) => id !== it.id)
      if (it !== to) await putItem(it)
    }
  }
  await putItem(to)
  await delItem(from)
}
/** 型號解讀：乾燥過濾器、膨脹閥看得懂時，就不再列「分數」那一條（重複） */
function decodedFor(f) {
  const list = decode(f.model, f.code, f.spec, f.label, f.text)
  return list.some((d) => ['drier', 'txv', 'coil'].includes(d.kind)) ? list.filter((d) => d.kind !== 'pipe') : list
}
/** 替代品：型號規則算出來的＋自己設定「可以互換」的；店裡有的標出來 */
function equivalentsFor(it, decoded, items) {
  const rule = decoded.flatMap((d) => d.equivalents.map((e) => ({ ...e, item: items.find((x) => x !== it && x.model && modelKey(x.model) === modelKey(e.model)) })))
  const linked = it ? items.filter((x) => x !== it && (it.equiv?.includes(x.id) || x.equiv?.includes(it.id))) : []
  return { rule, linked: linked.filter((x) => !rule.some((r) => r.item === x)) }
}

// ───────────────────────── 畫面狀態 ─────────────────────────
const state = { view: 'home', session: null, photoIndex: 0, focus: null, focusObj: null, busy: false, cancel: false, progress: null, refining: null, addMode: false, viewer: false, zoom: 2, itemFilter: 'all', itemId: null, lookup: { q: '', read: null, busy: false } }

function go(view, extra = {}) {
  Object.assign(state, { view, focus: null, focusObj: null, addMode: false, viewer: false }, extra)
  render()
  window.scrollTo({ top: 0 })
}
async function save() {
  if (state.session) await db.put(state.session)
}

// ───────────────────────── 畫面 ─────────────────────────
const chev = '<svg class="chev" width="10" height="17" viewBox="0 0 10 17" aria-hidden="true"><path d="M1.5 1.5 8 8.5l-6.5 7" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>'
const backBtn = (to = 'home', label = '盤點') => `<button class="back" data-go="${to}"><svg width="12" height="20" viewBox="0 0 12 20" aria-hidden="true"><path d="M10 2 2 10l8 8" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>${esc(label)}</button>`
/** 底部分頁列（像 iOS 的 Tab Bar）：盤點／品項／查型號 */
const ICON = {
  home: '<path d="M4 8.5 12 4l8 4.5v7L12 20l-8-4.5z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><path d="M4 8.5 12 13l8-4.5M12 13v7" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/>',
  items: '<rect x="4" y="4" width="16" height="16" rx="3.5" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M8 9h8M8 12.5h8M8 16h5" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>',
  lookup: '<circle cx="10.5" cy="10.5" r="6" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="m15 15 5 5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>',
}
const TABS = [
  { id: 'home', label: '盤點' },
  { id: 'items', label: '品項' },
  { id: 'lookup', label: '查型號' },
]
const tabBar = (active) =>
  `<nav class="tabbar" aria-label="主選單"><div class="inner">${TABS.map((t) => `<button data-go="${t.id}" ${t.id === active ? 'aria-current="page"' : ''}><svg width="26" height="26" viewBox="0 0 24 24" aria-hidden="true">${ICON[t.id]}</svg><span>${t.label}</span></button>`).join('')}</div></nav>`

async function viewHome() {
  const sessions = await db.all()
  const hasKey = !!ls.get(LS.key)
  return `
  <main class="app">
    <div class="nav">${syncReady() ? `<button class="btn small plain sync-pill ${state.syncState || ''}" data-action="sync-now">☁︎ ${esc(syncLabel())}</button>` : '<span></span>'}<button class="icon-btn" data-go="settings" aria-label="設定"><svg width="22" height="22" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 15.5A3.5 3.5 0 1 1 12 8.5a3.5 3.5 0 0 1 0 7Zm7.43-2.53a7.8 7.8 0 0 0 0-1.94l2.11-1.65a.5.5 0 0 0 .12-.64l-2-3.46a.5.5 0 0 0-.6-.22l-2.49 1a7.6 7.6 0 0 0-1.68-.98l-.38-2.65A.5.5 0 0 0 14 2h-4a.5.5 0 0 0-.5.42l-.37 2.65c-.61.25-1.17.58-1.69.98l-2.49-1a.5.5 0 0 0-.6.22l-2 3.46a.5.5 0 0 0 .12.64l2.11 1.65a7.8 7.8 0 0 0 0 1.94l-2.11 1.65a.5.5 0 0 0-.12.64l2 3.46a.5.5 0 0 0 .6.22l2.49-1c.52.4 1.08.73 1.69.98l.37 2.65c.04.24.25.42.5.42h4c.25 0 .46-.18.5-.42l.37-2.65c.61-.25 1.17-.58 1.68-.98l2.49 1a.5.5 0 0 0 .6-.22l2-3.46a.5.5 0 0 0-.12-.64l-2.1-1.65Z"/></svg></button></div>
    <h1 class="large-title">拍照盤點</h1>
    <p class="subtitle">拍貨架，AI 數品項；跟原圖對照，再用 ＋／－ 修正。</p>
    ${
      !canEdit()
        ? `<div class="hint-card"><b>你是檢視者（只能看）</b>：可以看大家的盤點紀錄、品項庫，也可以下載 Excel；不能盤點或修改。需要盤點請找管理員改成「編輯者」。</div>`
        : hasKey
        ? `<button class="hero-btn" data-action="new"><span class="hero-icon" aria-hidden="true">📷</span><span class="grow"><b>新盤點</b><br><span class="meta">拍一格貨架；有貼儲位標籤會自動填位置</span></span>${chev}</button>`
        : `<div class="hint-card stack"><div><b>要拍照辨識：</b>先到「設定」貼上你的免費 Gemini API Key（只會存在這台裝置）。<br>只想看手機盤點的結果（例如在電腦上）：到「設定 → 多人、多台同步」貼上連結碼就好，不用 Key。</div><button class="btn small" data-go="settings">去設定</button></div>`
    }
    ${(() => {
      if (!sessions.length) return ''
      const today = sessions.filter((s) => inRange(s, 'today'))
      const r = reportOf(today.length ? today : sessions)
      return `<button class="report-card" data-go="report">
        <span class="grow">
          <span class="report-card-kicker">總表與匯出・${today.length ? '今天' : '全部'}</span>
          <span class="report-card-nums"><b>${r.total}</b> 件・<b>${r.list.length}</b> 種・${r.sessions} 次盤點</span>
          <span class="meta">一次匯出 Excel、複製到 Google 試算表${ls.get(LS.sheet) ? '、同步' : ''}</span>
        </span>${chev}</button>`
    })()}
    <p class="section-title">盤點紀錄</p>
    ${
      sessions.length
        ? `<div class="group">${sessions
            .map((s) => {
              const failed = s.photos.length && s.photos.every((p) => p.status !== 'done')
              const meta = [byName(s) ? `${byName(s)} 盤・${fmtTime(s.createdAt)}` : fmtTime(s.createdAt), s.place && placeLabel(s.place), s.photos.length > 1 ? `${s.photos.length} 張照片` : '', s.linkedAt ? '' : '還沒按完成', s.syncedAt ? '已同步到試算表' : ''].filter(Boolean).join('・')
              return `<button class="row" data-open="${s.id}"><img src="${s.photos[0] ? urlOf(s.photos[0]) : ''}" alt="" style="width:52px;height:52px;border-radius:10px;object-fit:cover;background:var(--card-2)"><span class="grow"><span class="title">${esc(failed ? '沒有辨識成功（點進去再試一次）' : summaryOf(s))}</span><br><span class="meta">${esc(meta)}</span></span>${chev}</button>`
            })
            .join('')}</div>`
        : `<div class="empty"><div class="big">📦</div><p>還沒有盤點紀錄。<br>按上面「新盤點」，拍一層貨架試試看。</p></div>`
    }
  </main>
  ${tabBar('home')}`
}

function viewCapture() {
  const s = state.session
  return `
  <main class="app">
    <div class="nav">${backBtn('home', '盤點')}</div>
    <h1 class="large-title">新盤點</h1>
    ${
      s.by
        ? `<div class="group" style="margin-bottom:10px"><div class="row"><span style="width:96px" class="muted">盤點人</span><span class="grow"><b>${esc(s.by)}</b></span><span class="muted" style="font-size:14px">${canManage() ? '' : '由管理員設定'}</span>${canManage() ? '<button class="btn small plain" data-action="counter-session">換人</button>' : ''}</div></div>`
        : ''
    }
    <div class="group"><label class="row"><span style="width:96px" class="muted">位置</span><input class="inline" id="place" placeholder="${locations().length ? '點下面的儲位；拍到標籤也會自動填' : '例：A-01（拍到儲位標籤會自動填）'}" value="${esc(s.place)}" autocomplete="off"></label></div>
    ${
      locations().length
        ? `<div class="chips" role="group" aria-label="選儲位">${locations()
            .map((l) => `<button class="chip" data-place="${esc(l.code)}" aria-pressed="${canon(l.code) === canon(s.place)}">${esc(l.code)}${l.name ? `<small>${esc(l.name)}</small>` : ''}</button>`)
            .join('')}</div>`
        : `<p class="footnote">到「品項 → 儲位」建立代號、印標籤貼在貨架上：之後拍照會自動填位置，重盤同一格也會自動更新數量。</p>`
    }
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
  const doubts = doubtsOf(s)
  const doubtSet = new Set(doubts.map((d) => d.o))
  const boxesOf = (pi) =>
    s.photos[pi].objects
      .map((o, oi) => {
        const ref = colorOfRef.get(`${pi}:${oi}`)
        if (!ref) return ''
        const [y1, x1, y2, x2] = o.box
        const on = state.focusObj ? state.focusObj === o : state.focus === ref.key
        const doubt = doubtSet.has(o)
        return `<div class="box ${on ? 'on' : ''} ${doubt ? 'doubt' : ''}" style="--c:${ref.color};top:${y1 / 10}%;left:${x1 / 10}%;height:${(y2 - y1) / 10}%;width:${(x2 - x1) / 10}%" data-focus="${esc(ref.key)}" data-obj="${pi}:${oi}"><span class="tag">${ref.gi + 1}</span>${doubt ? '<span class="qmark" aria-label="要確認">?</span>' : ''}</div>`
      })
      .join('')
  const boxes = photo ? boxesOf(state.photoIndex) : ''
  const total = totalQty(s)
  const doubtsByKey = new Map()
  doubts.forEach((d) => doubtsByKey.set(keyOf(d.o), (doubtsByKey.get(keyOf(d.o)) || 0) + 1))
  const errors = s.photos.map((p, i) => (p.status !== 'done' ? `<div class="error-card">第 ${i + 1} 張沒辨識成功：${esc(p.error || '還沒辨識（被取消）')} <button class="btn small secondary" data-retry="${i}" style="margin-left:6px">再試一次</button>${p.errorDetail ? `<div style="margin-top:8px;font-size:12px;color:var(--text-2);word-break:break-all">錯誤代碼：${esc(p.errorDetail)}</div>` : ''}</div>` : '')).join('')
  const notes = s.photos.map((p, i) => (p.note ? `<p class="footnote">第 ${i + 1} 張 AI 備註：${esc(p.note)}</p>` : '')).join('')
  return `
  <main class="app">
    <div class="nav">${backBtn('home', '盤點')}<button class="btn small secondary edit-only" data-action="add">＋ 手動新增</button></div>
    <h1 class="large-title">${esc(s.place ? placeLabel(s.place) : '盤點結果')}</h1>
    ${
      byName(s) || syncReady()
        ? `<div class="by-line"><span class="avatar small" aria-hidden="true">${esc((byName(s) || '?').slice(0, 1))}</span><span class="grow"><b>盤點人：${esc(byName(s) || '沒有記錄')}</b><br><span class="meta">${fmtTime(s.createdAt)}${s.model ? `・${esc(s.model)}` : ''}</span></span>${canManage() && syncReady() ? '<button class="btn small plain" data-action="counter-session">更正</button>' : ''}</div>`
        : `<p class="subtitle">${fmtTime(s.createdAt)}${s.model ? `・${esc(s.model)}` : ''}</p>`
    }
    ${errors ? `<div class="stack">${errors}</div>` : ''}
    ${
      groups.length
        ? `<section class="summary" aria-label="盤點總結">
            <div class="sum-nums">
              <div><span class="sum-big">${total}</span><span class="sum-unit">件</span></div>
              <div class="sum-side"><b>${groups.length}</b> 種${s.photos.length > 1 ? `・${s.photos.length} 張照片` : ''}</div>
            </div>
            ${
              doubts.length
                ? `<button class="sum-doubt edit-only" data-action="review-doubts"><span class="sum-dot" aria-hidden="true">?</span><span class="grow"><b>${doubts.length} 個要確認</b><br><span class="meta">可能尺寸不同或 AI 沒把握；只看這幾個就好</span></span>${chev}</button>`
                : `<div class="sum-ok"><span aria-hidden="true">✓</span> 沒有需要確認的${state.refining && state.refining.session === s.id ? '（相似品還在比對）' : ''}</div>`
            }
          </section>`
        : ''
    }
    ${
      photo
        ? `<div class="photo-wrap ${state.focus || state.focusObj ? 'focus' : ''} ${state.addMode ? 'adding' : ''}" data-photo><img src="${urlOf(photo)}" alt="第 ${state.photoIndex + 1} 張照片">${boxes}
             <button class="photo-zoom" data-action="zoom" aria-label="放大看照片">⤢</button>
           </div>
           ${state.addMode ? '<div class="add-hint" role="status"><b>點照片上漏掉的那一個</b>，會在那裡加一個框 <button class="btn small plain" data-action="add-cancel">取消</button></div>' : ''}
           ${s.photos.length > 1 ? `<div class="photo-strip">${s.photos.map((p, i) => `<button class="${i === state.photoIndex ? 'on' : ''}" data-photo-index="${i}" aria-label="看第 ${i + 1} 張"><img src="${urlOf(p)}" alt=""></button>`).join('')}</div>` : ''}
           <div class="row-actions edit-only" style="margin-top:10px"><button class="btn small secondary" data-action="add-box" ${state.addMode ? 'disabled' : ''}>＋ 漏掉的，點照片補一個</button></div>
           <p class="footnote">點照片上的框：直接改成別的種類，改完自動跳下一個。點下面的清單：看那一種在哪裡；數量不對按 ＋／－。</p>`
        : ''
    }
    ${
      state.viewer && photo
        ? `<div class="viewer" role="dialog" aria-label="放大看照片">
            <div class="viewer-bar">
              <button class="btn small plain" data-action="zoom-close">完成</button>
              <span class="zoom-label">放大 ${state.zoom} 倍・可以上下左右滑</span>
              <button class="btn small secondary" data-action="zoom-out" ${state.zoom <= 1 ? 'disabled' : ''} aria-label="縮小">−</button>
              <button class="btn small secondary" data-action="zoom-in" ${state.zoom >= 4 ? 'disabled' : ''} aria-label="放大">＋</button>
            </div>
            <div class="viewer-scroll"><div class="photo-wrap viewer-photo ${state.focusObj ? 'focus' : ''}" style="width:${state.zoom * 100}%" data-photo><img src="${urlOf(photo)}" alt="">${boxes}</div></div>
          </div>`
        : ''
    }
    ${state.refining && state.refining.session === s.id ? `<div class="hint-card" role="status" style="margin-top:12px"><b>正在比對相似品…</b>（${state.refining.done}/${state.refining.total} 張）好了會自動更新，可以先看結果。</div>` : ''}
    ${notes}
    ${s.photos
      .map((p, i) => (p.trace?.length ? `<details class="trace"><summary>辨識過程${s.photos.length > 1 ? `（第 ${i + 1} 張）` : ''}：花了多久、用哪個模型</summary>${p.trace.map((l) => `<div>${esc(l)}</div>`).join('')}</details>` : ''))
      .join('')}
    <p class="section-title">品項（${groups.length}）</p>
    ${
      groups.length
        ? `<div class="group">${groups
            .map(
              (g, gi) => `
          <div class="item ${state.focus === g.key ? 'on' : ''}" data-item="${esc(g.key)}">
            <button class="swatch" style="--c:${g.color}" data-focus="${esc(g.key)}" aria-label="在照片上標出 ${esc(g.label)}">${gi + 1}</button>
            <button class="grow edit-btn" data-edit="${esc(g.key)}" aria-label="修改 ${esc(g.label)} 的名稱、品牌、型號、規格">
              <span class="name">${esc(g.label)}</span><span class="pencil" aria-hidden="true">✎</span>${g.manual ? '<span class="badge edit">手動</span>' : doubtsByKey.get(g.key) ? `<span class="badge low">${doubtsByKey.get(g.key)} 個要確認</span>` : ''}${g.edited && !g.manual ? '<span class="badge edit">已修正</span>' : ''}${itemsCache && !findItem(itemsCache, g) ? '<span class="badge ok">新品項</span>' : ''}
              <br><span class="spec">${esc(detailOf(g) || '點 ✎ 補品牌、型號、尺寸')}${!g.manual && g.count !== g.boxes ? `・照片裡 ${g.boxes} 個` : ''}</span>
            </button>
            <span class="stepper edit-only"><button data-step="-1" data-key="${esc(g.key)}" aria-label="減一">−</button><input inputmode="numeric" value="${g.count}" data-count="${esc(g.key)}" aria-label="${esc(g.label)} 數量"><button data-step="1" data-key="${esc(g.key)}" aria-label="加一">＋</button></span><span class="qty view-only"><b>${g.count}</b></span>
          </div>`,
            )
            .join('')}</div>`
        : `<div class="empty"><p>這次沒有找到商品。<br>可以重拍，或按右上「手動新增」。</p></div>`
    }
    <div class="row-actions edit-only" style="margin-top:22px">
      <button class="btn danger small" data-action="delete-session">刪除這次盤點</button>
    </div>
  </main>
  <div class="toolbar"><div class="inner"><button class="btn secondary" data-action="export">匯出</button><button class="btn edit-only" data-action="finish">完成・記進品項庫</button></div></div>`
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
    <div class="row-actions edit-only" style="margin-top:10px"><button class="btn small" data-action="save-catalog">儲存清單</button><button class="btn small secondary" data-action="reset-catalog">恢復預設</button></div>
    <p class="section-title">樣品照（${samples.length}）</p>
    ${
      samples.length
        ? `<div class="group">${samples
            .map(
              (s, i) =>
                `<div class="row"><img src="${sampleUrl(s)}" alt="" style="width:52px;height:52px;border-radius:10px;object-fit:contain;background:var(--card-2)"><span class="grow"><span class="title">${esc(s.label)}</span><br><span class="meta">${esc(detailOf(s) || '沒有填品牌、型號、尺寸')}${i >= MAX_SAMPLES ? '・超過 30 張，這張不會送' : ''}</span></span><button class="btn small danger edit-only" data-action="del-sample" data-id="${esc(s.id)}">刪除</button></div>`,
            )
            .join('')}</div>`
        : ''
    }
    <div class="row-actions edit-only" style="margin-top:10px"><label class="btn small secondary">📷 拍一張樣品<input type="file" accept="image/*" capture="environment" id="sample-cam" class="sr-only"></label></div>
    <p class="footnote">長得很像、只差尺寸的商品（例如不同分數的三通），每一種存一張樣品照，名稱和尺寸寫清楚。辨識時會一起送給 AI 比對（最多 30 張，新的優先）。<br>最快的存法：盤點結果裡先點那個框、再點一次 →「儲存，並存成樣品照」。</p>
    <p class="section-title">Google 試算表</p>
    <div class="stack">
      <input class="field" id="sheet-url" placeholder="貼上網頁應用程式網址（https://script.google.com/macros/s/…/exec）" value="${esc(ls.get(LS.sheet))}" autocomplete="off" spellcheck="false">
      <div class="row-actions"><button class="btn small" data-action="sheet-save">儲存並測試</button><button class="btn small secondary" data-action="sheet-copy">複製試算表程式碼</button></div>
      <label class="row" style="border-radius:12px;background:var(--card)"><span class="grow"><span class="title">按「完成」時自動同步</span><br><span class="meta">每次盤點完，自動寫進試算表</span></span><input type="checkbox" id="auto-sync" ${ls.get(LS.autoSync, '1') === '1' ? 'checked' : ''} style="width:22px;height:22px"></label>
    </div>
    <details class="steps"><summary>怎麼連結？（只要做一次，約 3 分鐘）</summary>
      <ol>
        <li>按上面「複製試算表程式碼」。</li>
        <li>電腦開一個新的 Google 試算表 → 上方「擴充功能」→「Apps Script」。</li>
        <li>把原本的內容全部刪掉，貼上剛剛複製的程式碼 → 存檔。</li>
        <li>右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」；執行身分選「我」，誰可以存取選「所有人」→ 部署。</li>
        <li>第一次會要你授權：選自己的帳號 →「進階」→「前往」→ 允許。</li>
        <li>複製「網頁應用程式網址」，貼到上面的格子 → 按「儲存並測試」。</li>
      </ol>
      <p>試算表會自動建立五張：「總覽」（給主管：品項、件數、該叫貨、盤虧盤盈、儲位盤點進度、圖表）、「庫存」（每個品項實盤、帳面、差異、在哪裡、誰盤的）、「儲位庫存」（每一格放了什麼）、「盤差報告」（實盤≠帳面的，後面三欄給主管填原因、處理方式、確認，重新同步會保留）、「盤點紀錄」（原始資料）。每次同步自動更新，手動改會先跳警告。資料只會寫進你自己的試算表。</p>
      <p><b>以前連結過的：</b>要有「品項庫」工作表和多台同步，請重新複製程式碼貼上 → 存檔 →「部署」→「管理部署作業」→ ✎ 編輯 → 版本選「新版本」→ 部署（網址不變；會再問一次授權，因為要存到你的雲端硬碟）。</p>
    </details>
    <p class="section-title">共用（跟 Google 雲端硬碟一樣）</p>
    ${
      syncReady()
        ? `<div class="group">
            <div class="row"><span class="avatar" aria-hidden="true">${esc((ls.get(LS.memberName) || '我').slice(0, 1))}</span><span class="grow"><span class="title">${esc(ls.get(LS.memberName) || '我')}（這台）</span><br><span class="meta">${ROLE_LABEL[myRole()]}・${state.syncState === 'error' && state.syncError ? esc(state.syncError) : esc(syncLabel())}</span></span><button class="btn small" data-action="sync-now">立即同步</button></div>
            <div class="row"><span class="grow"><span class="title">這台的盤點人</span><br><span class="meta">新盤點會自動記成這個人，不用每次選${canManage() ? '' : '；由擁有者或管理員設定'}</span></span>${
              canManage() ? `<button class="btn small secondary" data-action="counter-device">${esc(currentCounter()?.name || '選擇')}</button>` : `<b>${esc(currentCounter()?.name || '—')}</b>`
            }</div>
            ${
              canManage()
                ? `<button class="row" data-action="share-open"><span class="grow"><span class="title" style="color:var(--tint)">共用設定</span><br><span class="meta">邀請同事、改權限（只能看／可以改）、移除離職的人</span></span>${chev}</button>`
                : `<div class="row muted">${ROLE_DESC[myRole()]}。要加人或改權限，請找擁有者或管理員。</div>`
            }
            <button class="row" data-action="sync-leave"><span class="grow"><span class="title" style="color:var(--red)">這台退出並清除資料</span><br><span class="meta">交還手機、換手機時用；雲端的資料不會刪</span></span>${chev}</button>
          </div>`
        : ls.get(LS.sheet)
          ? '<button class="btn block" data-action="sync-start">開啟多人同步（我是擁有者）</button><p class="footnote">開啟後到「共用設定」邀請同事：每個人一組自己的連結碼，可以設「只能看」或「可以改」。</p>'
          : '<div class="group"><div class="row muted">擁有者：先完成上面的 Google 試算表連結，再回來開啟同步。<br>被邀請的人：直接在下面貼上收到的連結碼。</div></div>'
    }
    ${
      syncReady()
        ? ''
        : `<details class="steps" open><summary>我收到連結碼了</summary>
      <div class="stack" style="margin-top:8px">
        <input class="field" id="link-code" placeholder="貼上連結碼（https://script.google.com/…#k=…）" autocomplete="off" spellcheck="false">
        <button class="btn small" data-action="sync-link">加入</button>
      </div>
    </details>`
    }
    <details class="steps"><summary>資料安全嗎？（公司資產）</summary>
      <ol>
        <li><b>資料放在哪：</b>只在擁有者的 Google 雲端硬碟「拍照盤點同步資料」資料夾和試算表。可以先用個人帳號，之後在「共用設定 → 搬到另一個 Google 帳號」搬到公司帳號（大家自動跟過去）。GitHub 上只有程式，沒有任何盤點資料。</li>
        <li><b>誰讀得到：</b>只有共用名單裡的人。每個人一組自己的連結碼（亂數，猜不到；雲端只存雜湊值），傳輸全程加密（HTTPS）。部署時選的「所有人」只代表可以呼叫網址，沒有連結碼一律拒絕。</li>
        <li><b>權限：</b>擁有者、管理員（可以邀請／移除人）、編輯者（可以盤點、修改）、檢視者（只能看）。改權限馬上生效，不用換連結碼。</li>
        <li><b>有人離職：</b>在「共用設定」把他「移除權限」就好，其他人不用改；他的手機下次連線時，App 裡的公司資料會自動清除。</li>
        <li><b>手機上也有一份：</b>每台裝置會存一份方便離線看；手機請設螢幕鎖。</li>
        <li><b>拍照辨識：</b>照片會送到 Google Gemini 分析。免費版的條款寫明：Google 可以用送去的內容改善產品，也可能有人工審閱。擔心的話，到 Google AI Studio 開啟付費（照用量計費），付費版不會拿去改善產品。照片裡不要拍到價格單、客戶資料。</li>
      </ol>
    </details>
    <p class="section-title">連線測試</p>
    <div class="stack"><button class="btn small secondary" data-action="diagnose">測試連線</button><div id="diag"></div></div>
    <p class="footnote">辨識一直失敗時按這個，把結果截圖給我看。</p>
    <p class="section-title">資料</p>
    <div class="row-actions edit-only"><button class="btn small danger" data-action="clear-all">刪除全部盤點紀錄</button></div>
    <p class="footnote">紀錄（含照片）只存在這支手機的瀏覽器裡；要留底請用「匯出」。品項庫請到「品項 → ⋯ → 備份品項庫」。</p>
    <p class="footnote" style="margin-top:18px;text-align:center">拍照盤點 版本 ${VERSION}</p>
    <div class="row-actions" style="justify-content:center"><button class="btn small secondary" data-action="force-update">檢查更新</button></div>
    <p class="footnote" style="text-align:center">有新版會自動更新；不放心就按這裡。盤點紀錄、樣品照、API Key 都不會被刪。</p>
  </main>`
}

// ───────────────────────── 品項庫畫面 ─────────────────────────
const ITEM_FILTERS = [
  { id: 'all', label: '全部', test: () => true },
  { id: 'order', label: '叫貨', test: needsOrder },
  { id: 'diff', label: '差異', test: (it) => (diffOf(it) ?? 0) !== 0 },
  { id: 'new', label: '新的', test: (it) => it.status === 'new' },
]
const itemBadges = (it) => {
  const d = diffOf(it)
  return `${it.status === 'new' ? '<span class="badge ok">新的</span>' : ''}${needsOrder(it) ? '<span class="badge low">該叫貨</span>' : ''}${d ? `<span class="badge ${d < 0 ? 'bad' : 'edit'}">${d > 0 ? '+' : ''}${d}</span>` : ''}`
}
const itemRow = (it) => {
  const live = liveStock(it).map(([, st]) => st)
  const places = live
    .sort((a, b) => b.count - a.count)
    .map((st) => `${st.place || '沒填位置'} ${st.count}`)
    .join('、')
  const search = canon([it.no, it.label, it.brand, it.model, it.spec, ...live.map((st) => st.place)].join(' '))
  return `<button class="row item-row" data-item-open="${it.id}" data-search="${esc(search)}">
    ${itemThumb(it)}
    <span class="grow"><span class="title">${esc(itemTitle(it))}</span>${itemBadges(it)}<br><span class="meta">${esc([it.no, it.brand, it.model].filter(Boolean).join('・'))}${places ? `<br>${esc(places)}` : ''}</span></span>
    <span class="qty"><b>${onHand(it)}</b>${it.book != null ? `<small>帳面 ${it.book}</small>` : ''}</span>${chev}</button>`
}

async function viewItems() {
  const items = await itemsAll(true)
  const f = ITEM_FILTERS.find((x) => x.id === state.itemFilter) || ITEM_FILTERS[0]
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  const list = items.filter(f.test).sort((a, b) => cmp(a.label, b.label) || cmp(a.spec, b.spec) || cmp(a.model, b.model))
  const byLabel = new Map()
  for (const it of list) byLabel.set(it.label, [...(byLabel.get(it.label) || []), it])
  return `
  <main class="app">
    <div class="nav"><button class="btn small plain" data-go="locations">儲位</button><span class="nav-right"><button class="icon-btn" data-action="items-more" aria-label="匯入、匯出、備份">⋯</button><button class="btn small edit-only" data-action="item-add">＋ 新增</button></span></div>
    <h1 class="large-title">品項庫</h1>
    <p class="subtitle">${items.length ? `${items.length} 種商品・${locations().length} 個儲位。盤點按「完成」就會自動更新。` : '盤點按「完成」，數到的東西就會自動記進來。'}</p>
    ${
      items.length
        ? `<div class="seg" role="tablist" aria-label="篩選">${ITEM_FILTERS.map((x) => `<button role="tab" aria-selected="${x.id === f.id}" data-item-filter="${x.id}">${x.label} ${items.filter(x.test).length}</button>`).join('')}</div>
           ${f.id === 'all' && items.some((it) => it.status === 'new') ? `<button class="sum-doubt tip edit-only" data-item-filter="new"><span class="sum-dot" aria-hidden="true">!</span><span class="grow"><b>${items.filter((it) => it.status === 'new').length} 個新的品項，請確認名稱</b><br><span class="meta">盤點時 AI 自動建立的；名稱對就按「確認」，重複的就合併</span></span>${chev}</button>` : ''}
           ${f.id === 'order' && list.length ? `<button class="btn secondary block" data-action="order-copy" style="margin-bottom:6px">複製叫貨清單（貼到 LINE）</button>` : ''}
           ${f.id === 'order' && !list.length ? '' : '<input class="field search" id="item-search" type="search" placeholder="搜尋品名、型號、料號、儲位" autocomplete="off" enterkeyhint="search">'}
           ${
             list.length
               ? [...byLabel]
                   .map(([label, its]) => `<section class="item-sec"><p class="section-title">${esc(label)}（${its.length}）</p><div class="group">${its.map(itemRow).join('')}</div></section>`)
                   .join('')
               : f.id === 'order'
                 ? `<div class="hint-card stack" style="margin-top:12px">
                      <div><b>還沒有要叫貨的。</b></div>
                      <div>先告訴 App 每一種「<b>剩幾個就要叫貨</b>」：數量剩這麼多（或更少）時，就會出現在這裡。${items.some((it) => it.safety != null) ? '' : '<br>目前每一種都還沒設定。'}</div>
                      <button class="btn small edit-only" data-action="safety-pick">設定「剩幾個就要叫貨」</button>
                    </div>`
                 : `<div class="empty"><p>${f.id === 'diff' ? '實盤跟帳面都一樣。<br>（要先在品項裡設定「帳面數」才會比對）' : '沒有新的品項，都確認過了。'}</p></div>`
           }
           <p class="empty" id="search-empty" hidden>找不到。可以到「查型號」用型號找替代品。</p>
           <details class="steps" style="margin-top:18px"><summary>品項庫怎麼用？</summary>
             <ol>
               <li><b>記進來：</b>盤點完按「完成・記進品項庫」，數到的東西會自動記進來。也可以按右上「＋ 新增」，或到「查型號」拍標籤加入。</li>
               <li><b>確認名稱：</b>標「新的」是 AI 自動建立的，點進去看名稱對不對，對就按「確認」。</li>
               <li><b>叫貨提醒：</b>點進一個品項，設「剩幾個就要叫貨」；數量少於這個數字，就會出現在上面的「叫貨」。</li>
               <li><b>帳面數：</b>點進品項按「設定帳面數 → 用實盤數」當起點；之後進貨按「＋ 進貨」、賣掉按「－ 賣出」。下次盤點數量跟帳面不一樣，就會出現在「差異」。</li>
             </ol>
           </details>`
        : `<div class="hint-card stack">
             <div><b>品項庫會自己長出來</b>：不用先建好。每次盤點按「完成・記進品項庫」，數到的每一種都會變成一筆（有料號、在哪裡、幾個）。</div>
             <div>也可以：</div>
             <div class="row-actions"><button class="btn small" data-go="lookup">📷 拍型號加進來</button><button class="btn small secondary" data-action="import">貼上 Excel 清單</button><button class="btn small secondary" data-go="locations">建立儲位</button></div>
           </div>`
    }
  </main>
  ${tabBar('items')}`
}

async function viewItem() {
  const items = await itemsAll()
  const it = items.find((x) => x.id === state.itemId)
  if (!it) return viewItems()
  const total = onHand(it)
  const last = lastCounted(it)
  const d = diffOf(it)
  const decoded = decodedFor(it)
  const eq = equivalentsFor(it, decoded, items)
  const stock = liveStock(it).sort((a, b) => b[1].at - a[1].at)
  const links = linksFor(it, decoded)
  return `
  <main class="app">
    <div class="nav">${backBtn('items', '品項')}<button class="btn small secondary edit-only" data-action="item-edit">編輯</button></div>
    <div class="item-head">${itemThumb(it, 64)}<div class="grow"><h1 class="large-title">${esc(itemTitle(it))}</h1><p class="subtitle">${esc([it.no, it.brand, it.model].filter(Boolean).join('・'))}</p></div></div>
    ${
      it.status === 'new'
        ? `<div class="hint-card stack edit-only"><div><b>盤點時自動建立的。</b>名稱、尺寸對嗎？跟別的品項重複就合併。</div><div class="row-actions"><button class="btn small" data-action="item-confirm">✓ 對，確認</button><button class="btn small secondary" data-action="item-edit">修改</button><button class="btn small secondary" data-action="item-merge">合併到…</button></div></div>`
        : ''
    }
    <section class="summary">
      <div class="stock-nums">
        <div><span class="stock-label">實盤</span><span class="stock-big">${total}</span><span class="stock-sub">${last ? `最近 ${fmtTime(last)}` : '還沒盤過'}</span></div>
        <div><span class="stock-label">帳面</span><span class="stock-big">${it.book ?? '—'}</span><span class="stock-sub">${it.book == null ? '還沒設定' : '進貨加、賣出減'}</span></div>
        <div><span class="stock-label">差異</span><span class="stock-big ${d < 0 ? 'neg' : d > 0 ? 'pos' : ''}">${d == null ? '—' : `${d > 0 ? '+' : ''}${d}`}</span><span class="stock-sub">${d == null ? '設定帳面數才比' : d > 0 ? '盤盈（多了）' : d < 0 ? '盤虧（少了）' : '一樣'}</span></div>
      </div>
      <div class="row-actions edit-only" style="margin-top:14px"><button class="btn small secondary" data-action="move-in">＋ 進貨</button><button class="btn small secondary" data-action="move-out">－ 賣出</button><button class="btn small secondary" data-action="book-set">設定帳面數</button></div>
    </section>
    <p class="section-title">叫貨提醒</p>
    <div class="group"><div class="row"><span class="grow"><span class="title">剩幾個就要叫貨</span><br><span class="meta">${needsOrder(it) ? `⚠️ 現在大概剩 ${expected(it)} 個，該叫貨了` : it.safety == null ? '按 ＋ 設一個數字，例如 3：剩 3 個以下就會出現在「叫貨」' : `剩 ${it.safety} 個以下，就會出現在「叫貨」`}</span></span><span class="stepper edit-only"><button data-safety="-1" aria-label="減一">−</button><input id="safety" inputmode="numeric" value="${it.safety ?? ''}" placeholder="—" aria-label="剩幾個就要叫貨"><button data-safety="1" aria-label="加一">＋</button></span><span class="qty view-only"><b>${it.safety ?? '—'}</b></span></div></div>
    <p class="section-title">在哪裡（${stock.length} 個位置）</p>
    ${
      stock.length
        ? `<div class="group">${stock.map(([k, st]) => `<div class="row"><span class="grow"><span class="title">${esc(stockPlace(st))}</span><br><span class="meta">${fmtTime(st.at)} 盤點</span></span><span class="qty"><b>${st.count}</b></span><button class="icon-btn small edit-only" data-stock-del="${esc(k)}" aria-label="拿掉這個位置的數量">×</button></div>`).join('')}</div>`
        : '<div class="group"><div class="row muted">還沒盤點過。盤點時按「完成」就會記在這裡。</div></div>'
    }
    ${
      decoded.length
        ? `<p class="section-title">型號解讀</p><div class="group">${decoded.map((x) => `<div class="row decode"><span class="grow"><span class="title">${esc(x.title)}</span>${x.facts.map((t) => `<br><span class="meta">・${esc(t)}</span>`).join('')}</span></div>`).join('')}</div>`
        : ''
    }
    <p class="section-title">替代品（可以互換）</p>
    <div class="group">
      ${eq.rule.map((e) => equivRow(e)).join('')}
      ${eq.linked.map((x) => equivRow({ brand: x.brand, model: x.model || itemTitle(x), item: x, linked: true })).join('')}
      <button class="row edit-only" data-action="equiv-add"><span class="swatch" style="--c:var(--tint);pointer-events:none">＋</span><span class="grow"><span class="title" style="color:var(--tint)">加一個可以互換的品項</span><br><span class="meta">例如客人常問的別牌同規格</span></span></button>
    </div>
    ${decoded.find((x) => x.note)?.note ? `<p class="footnote">${esc(decoded.find((x) => x.note).note)}</p>` : ''}
    ${links.length ? `<p class="section-title">查原廠資料</p><div class="group">${links.map((l) => `<a class="row" href="${esc(l.url)}" target="_blank" rel="noreferrer"><span class="grow">${esc(l.title)}</span>${chev}</a>`).join('')}</div>` : ''}
    ${
      it.moves?.length
        ? `<p class="section-title">進出紀錄</p><div class="group">${[...it.moves]
            .reverse()
            .slice(0, 20)
            .map((m) => `<div class="row"><span class="grow">${m.kind === 'in' ? '進貨' : m.kind === 'out' ? '賣出' : '設定帳面數'}<br><span class="meta">${fmtTime(m.at)}</span></span><span class="qty"><b>${m.kind === 'in' ? '+' : m.kind === 'out' ? '−' : '＝'}${m.qty}</b></span></div>`)
            .join('')}</div>`
        : ''
    }
    <div class="row-actions edit-only" style="margin-top:22px"><button class="btn small secondary" data-action="item-merge">合併到另一個品項</button><button class="btn small danger" data-action="item-delete">刪除品項</button></div>
    <p class="footnote">以前的寫法（AI 認過的名稱）：${esc((it.aliases || []).length)} 種，以後辨識到都會算進這一項。</p>
  </main>`
}
function equivRow(e) {
  const x = e.item
  return x
    ? `<button class="row" data-item-open="${x.id}"><span class="grow"><span class="title">${esc([e.brand, e.model].filter(Boolean).join(' '))}</span><br><span class="meta in-stock">店裡有 ${onHand(x)}・${esc(x.no)}${e.linked ? '・自己設定的' : ''}</span></span>${chev}</button>`
    : `<div class="row"><span class="grow"><span class="title">${esc([e.brand, e.model].filter(Boolean).join(' '))}</span><br><span class="meta">店裡還沒有這一項</span></span></div>`
}

// ───────────────────────── 查型號：客人拿零件來問 ─────────────────────────
const EXAMPLES = ['DML 083S', 'ADK-163', 'TES 2', 'KP 15', 'EVR 6', '4×11×330', '1吋1分', 'R404A']
function lookupResults() {
  const { q, read } = state.lookup
  const items = itemsCache || []
  const fields = read || { model: q, label: q }
  const decoded = decodedFor(read ? read : { model: q })
  const main = decoded.find((x) => x.model)
  const cq = canon(q)
  const matches =
    cq.length >= 2
      ? items
          .filter((it) => {
            const hay = canon([it.no, it.label, it.brand, it.model, it.spec].join(' '))
            return hay.includes(cq) || (main && it.model && modelKey(it.model) === modelKey(main.model)) || (read && findItem([it], read))
          })
          .slice(0, 20)
      : []
  const eq = equivalentsFor(matches[0] || null, decoded, items)
  const links = linksFor({ brand: read?.brand || main?.brand || '', model: read?.model || main?.model || q, label: read?.label || '' }, decoded)
  if (!q && !read)
    return `<p class="section-title">可以查（點一個試試看）</p>
      <div class="chips">${EXAMPLES.map((x) => `<button class="chip" data-example="${esc(x)}">${esc(x)}</button>`).join('')}</div>
      <p class="footnote">看得懂：各牌乾燥過濾器（DML、DCL、ADK、EK、FD、C-）、Danfoss 膨脹閥（TEX、TES、TEN）、KP 壓力開關、EVR 電磁閥、散熱器排×支×鏡面、銅管分數、冷媒、Danfoss 訂購碼。看不懂的型號也會找店裡有沒有，並附原廠搜尋連結。</p>`
  const exact = read ? items.find((it) => findItem([it], read)) : null
  return `
    ${
      read
        ? `<section class="summary read-card">
            <div class="read-head">${read.url ? `<img src="${read.url}" alt="拍到的標籤">` : ''}<div class="grow"><span class="stock-label">AI 讀到</span><b>${esc(read.label || '（看不出品名）')}</b><br><span class="meta">${esc([read.brand, read.model, read.spec].filter(Boolean).join('・') || '看不出型號')}</span>${read.code ? `<br><span class="meta">訂購碼 ${esc(read.code)}</span>` : ''}</div></div>
            ${read.text ? `<details class="trace"><summary>標籤上的字</summary>${esc(read.text)}</details>` : ''}
            <div class="row-actions" style="margin-top:12px">${exact ? `<button class="btn small" data-item-open="${exact.id}">打開 ${esc(exact.no)}（店裡有 ${onHand(exact)}）</button>` : '<button class="btn small edit-only" data-action="read-add">加入品項庫</button>'}<button class="btn small secondary" data-action="read-clear">清除</button></div>
          </section>`
        : ''
    }
    ${decoded.length ? `<p class="section-title">這是什麼</p><div class="group">${decoded.map((x) => `<div class="row decode"><span class="grow"><span class="title">${esc(x.title)}</span>${x.facts.map((t) => `<br><span class="meta">・${esc(t)}</span>`).join('')}</span></div>`).join('')}</div>` : ''}
    <p class="section-title">店裡有的（${matches.length}）</p>
    ${matches.length ? `<div class="group">${matches.map(itemRow).join('')}</div>` : `<div class="group"><div class="row muted">品項庫裡沒有${fields.model || fields.label ? `「${esc(read ? read.model || read.label : q)}」` : ''}。</div></div>`}
    ${
      eq.rule.length || eq.linked.length
        ? `<p class="section-title">替代品（同規格，可以互換）</p><div class="group">${eq.rule.map(equivRow).join('')}${eq.linked.map((x) => equivRow({ brand: x.brand, model: x.model || itemTitle(x), item: x, linked: true })).join('')}</div>${decoded.find((x) => x.note)?.note ? `<p class="footnote">${esc(decoded.find((x) => x.note).note)}</p>` : ''}`
        : ''
    }
    ${links.length ? `<p class="section-title">查原廠資料</p><div class="group">${links.map((l) => `<a class="row" href="${esc(l.url)}" target="_blank" rel="noreferrer"><span class="grow">${esc(l.title)}</span>${chev}</a>`).join('')}</div>` : ''}`
}
async function viewLookup() {
  await itemsAll()
  const { q, busy } = state.lookup
  return `
  <main class="app">
    <div class="nav"><span></span></div>
    <h1 class="large-title">查型號</h1>
    <p class="subtitle">客人拿零件或型號來問：拍標籤或打型號，馬上看是什麼、店裡有沒有、可以用什麼替代。</p>
    <label class="hero-btn ${busy ? 'busy' : ''}" ${busy ? 'aria-disabled="true"' : ''}><span class="hero-icon" aria-hidden="true">${busy ? '<span class="spinner small"></span>' : '📷'}</span><span class="grow"><b>${busy ? 'AI 讀標籤中…' : '拍型號標籤'}</b><br><span class="meta">${busy ? '大約 5～15 秒' : '外盒、貼紙、機器上的型號牌、零件上的刻字都可以'}</span></span><input type="file" accept="image/*" capture="environment" id="label-cam" class="sr-only" ${busy || !ls.get(LS.key) ? 'disabled' : ''}></label>
    ${ls.get(LS.key) ? '' : '<p class="footnote">拍標籤要先到「設定」貼上 API Key；打型號查詢不用。</p>'}
    <input class="field search" id="lookup-q" type="search" placeholder="或打型號：DML 083S、TES 2、4×11×330" value="${esc(q)}" autocomplete="off" enterkeyhint="search" spellcheck="false">
    <div id="lookup-results">${lookupResults()}</div>
  </main>
  ${tabBar('lookup')}`
}

// ───────────────────────── 儲位畫面 ─────────────────────────
async function viewLocations() {
  const items = await itemsAll()
  const locs = locations()
  const inLoc = (it, code) => liveStock(it).filter(([, st]) => canon(st.place) === canon(code))
  const here = (code) => items.filter((it) => inLoc(it, code).length)
  return `
  <main class="app">
    <div class="nav">${backBtn('items', '品項')}</div>
    <h1 class="large-title">儲位</h1>
    <p class="subtitle">每一格貨架一個代號（例如 A-01）。印成標籤貼在貨架上：拍照時 AI 看到標籤，就會自動填位置；重盤同一格，數量自動更新。</p>
    ${
      locs.length
        ? `<div class="group">${locs
            .map((l, i) => {
              const its = here(l.code)
              const qty = its.reduce((n, it) => n + inLoc(it, l.code).reduce((m, [, st]) => m + st.count, 0), 0)
              return `<button class="row" data-loc-edit="${i}"><span class="loc-code">${esc(l.code)}</span><span class="grow"><span class="title">${esc(l.name || '（沒有說明）')}</span><br><span class="meta">${its.length ? `${its.length} 種・${qty} 件` : '還沒盤點'}</span></span>${chev}</button>`
            })
            .join('')}</div>`
        : '<div class="group"><div class="row muted">還沒有儲位。按下面「新增儲位」，可以一次建立一整排。</div></div>'
    }
    <div class="row-actions" style="margin-top:14px"><button class="btn edit-only" style="flex:1" data-action="loc-add">＋ 新增儲位</button><button class="btn secondary" style="flex:1" data-action="loc-print" ${locs.length ? '' : 'disabled'}>列印標籤</button></div>
    <details class="steps"><summary>怎麼編號比較好？</summary>
      <ol><li>字母＝第幾排貨架（A、B、C…），數字＝第幾層（由上往下 01、02…）。例：B-03＝B 排第 3 層。</li><li>標籤印出來剪下，貼在每一層的正中間、正面朝外；拍照時把標籤一起拍進去。</li><li>同一格要一次拍完（可以拍好幾張）。重盤同一格，會以新的那次為準（這次沒拍到的，就從這一格拿掉）；不同格的數量會加起來變成「實盤」。</li></ol>
    </details>
  </main>`
}

async function render() {
  const html =
    state.view === 'home'
      ? await viewHome()
      : state.view === 'capture'
        ? viewCapture()
        : state.view === 'analyzing'
          ? viewAnalyzing()
          : state.view === 'review'
            ? viewReview()
            : state.view === 'report'
              ? await viewReport()
              : state.view === 'items'
                ? await viewItems()
                : state.view === 'item'
                  ? await viewItem()
                  : state.view === 'lookup'
                    ? await viewLookup()
                    : state.view === 'locations'
                      ? await viewLocations()
                      : await viewSettings()
  // 放大看照片時，重畫畫面不要讓位置跳回左上角
  const vs = document.querySelector('.viewer-scroll')
  const keep = vs ? { x: vs.scrollLeft / Math.max(1, vs.scrollWidth), y: vs.scrollTop / Math.max(1, vs.scrollHeight) } : null
  $app.innerHTML = html
  const nv = document.querySelector('.viewer-scroll')
  if (keep && nv) {
    nv.scrollLeft = keep.x * nv.scrollWidth
    nv.scrollTop = keep.y * nv.scrollHeight
  }
  document.body.classList.toggle('no-scroll', !!document.querySelector('.viewer'))
  // 檢視者：所有修改用的按鈕藏起來（CSS：body.read-only .edit-only）
  document.body.classList.toggle('read-only', !canEdit())
  bindInputs()
}

// ───────────────────────── 底部面板（Sheet）：點旁邊暗處或按 Esc 就關 ─────────────────────────
/** onDismiss：使用者點旁邊暗處或按 Esc 關掉時才呼叫（程式自己 close() 不會） */
function sheet(html, onMount, onDismiss) {
  const back = document.createElement('div')
  back.className = 'sheet-backdrop'
  back.innerHTML = `<div class="sheet" role="dialog" aria-modal="true"><div class="grabber"></div>${html}</div>`
  const close = () => {
    back.remove()
    document.removeEventListener('keydown', onKey)
  }
  const dismiss = () => {
    close()
    onDismiss?.()
  }
  const onKey = (e) => e.key === 'Escape' && dismiss()
  back.addEventListener('click', (e) => e.target === back && dismiss())
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

/**
 * 點照片上的框：「這一個是哪一種？」按一下就改，自動跳到下一個框；一路點下去就分完。
 * 順序是由上到下、由左到右；也可以按 ‹ › 自己跳。
 */
/** 補框：在點的位置加一個框（大小取這張照片其他框的中間值），標成手動補的，馬上問它是哪一種 */
async function addBoxAt(wrap, e) {
  const s = state.session
  const pi = state.photoIndex
  const photo = s.photos[pi]
  const rect = wrap.getBoundingClientRect()
  const x = ((e.clientX - rect.left) / rect.width) * 1000
  const y = ((e.clientY - rect.top) / rect.height) * 1000
  const ws = photo.objects.map((o) => o.box[3] - o.box[1]).sort((a, b) => a - b)
  const hs = photo.objects.map((o) => o.box[2] - o.box[0]).sort((a, b) => a - b)
  const w = ws[Math.floor(ws.length / 2)] || 120
  const h = hs[Math.floor(hs.length / 2)] || 120
  const clamp = (v) => Math.min(1000, Math.max(0, Math.round(v)))
  // 預設跟這張照片最多的那一種一樣，等一下可以改
  const groups = groupsOf(s).filter((g) => !g.manual && g.refs.some((r) => r.pi === pi))
  const base = [...groups].sort((a, b) => b.boxes - a.boxes)[0] || { label: '商品', brand: '', model: '', spec: '' }
  const o = { label: base.label, brand: base.brand, model: base.model, spec: base.spec, box: [clamp(y - h / 2), clamp(x - w / 2), clamp(y + h / 2), clamp(x + w / 2)], confidence: 1, edited: true, added: true }
  photo.objects.push(o)
  state.addMode = false
  await save()
  render()
  toast('補了一個；選它是哪一種')
  quickSheet([{ pi, o, reason: '你剛剛補的' }], 0)
}

/** 一張照片的所有框，照閱讀順序（由上到下、由左到右） */
function photoEntries(pi) {
  const photo = state.session.photos[pi]
  const center = (o) => ({ y: (o.box[0] + o.box[2]) / 2, x: (o.box[1] + o.box[3]) / 2 })
  return [...photo.objects].sort((a, b) => Math.round(center(a).y / 120) - Math.round(center(b).y / 120) || center(a).x - center(b).x).map((o) => ({ pi, o }))
}

/**
 * entries：要一個一個看的框（{ pi, o, reason }）；doubt＝「只看要確認的」模式（多一個「一樣，沒問題」）。
 * 跨照片時會自動切到那一張，後面的大圖跟著換。
 */
function quickSheet(entries, start = 0, { doubt = false } = {}) {
  const s = state.session
  if (!entries.length) return
  // 用物件本身記順序（刪掉框時索引會變）
  const order = [...entries]
  let cur = Math.max(0, Math.min(start, order.length - 1))
  let previewUrl = ''
  let formOpen = false
  sheet('<div id="q-body"></div>', (el, close) => {
    const body = el.querySelector('#q-body')
    const finish = () => {
      state.focusObj = null
      if (previewUrl) URL.revokeObjectURL(previewUrl)
      close()
      render()
    }
    const go = (i) => {
      if (i < 0) return
      if (i >= order.length) {
        if (doubt) toast('要確認的都看完了')
        return finish()
      }
      cur = i
      formOpen = false
      draw()
    }
    const assign = async (fields) => {
      const { pi, o } = order[cur]
      const oi = s.photos[pi].objects.indexOf(o)
      if (oi < 0) return go(cur + 1)
      moveObjects(s, [{ pi, oi }], fields)
      await save()
      render()
      if (cur + 1 >= order.length) {
        toast(doubt ? '要確認的都看完了' : '已經是最後一個了')
        return finish()
      }
      go(cur + 1)
    }
    const draw = async () => {
      const { pi, o, reason } = order[cur]
      const photo = s.photos[pi]
      state.focusObj = o
      if (state.photoIndex !== pi) state.photoIndex = pi
      render()
      const groups = groupsOf(s).filter((g) => !g.manual)
      const mine = groups.find((g) => g.key === keyOf(o))
      body.innerHTML = `
        <div style="display:flex;align-items:center;justify-content:space-between;gap:8px">
          <h2 class="sheet-title" style="margin:0">${doubt ? '這一個一樣嗎？' : '這一個是哪一種？'}</h2>
          <span class="muted" style="font-size:15px">${doubt ? '要確認的' : ''}第 ${cur + 1} / ${order.length} 個${s.photos.length > 1 ? `・第 ${pi + 1} 張照片` : ''}</span>
        </div>
        <img id="q-img" alt="這一個框的放大圖" style="display:block;width:100%;height:120px;object-fit:contain;margin:10px 0 8px;border-radius:12px;background:var(--card-2)">
        ${reason ? `<p class="doubt-reason">為什麼要看：${esc(reason)}</p>` : ''}
        <p class="sheet-sub" style="margin:0 0 10px">目前：${mine ? `<b style="color:${mine.color}">${groups.indexOf(mine) + 1}</b> ${esc(o.label)}${detailOf(o) ? `・${esc(detailOf(o))}` : ''}` : esc(o.label)}</p>
        ${doubt ? `<button class="btn block" id="q-same" style="margin-bottom:12px">✓ 一樣，就是「${esc(mine ? `${mine.label}${mine.spec ? `・${mine.spec}` : ''}` : o.label)}」</button><p class="sheet-sub" style="margin:0 0 8px">不一樣的話，選它是哪一種：</p>` : ''}
        <div class="group">
          ${groups
            .map(
              (g, gi) => `<button class="row" data-q="${gi}" ${g === mine ? 'aria-current="true"' : ''}>
                <span class="swatch" style="--c:${g.color};pointer-events:none">${gi + 1}</span>
                <span class="grow"><span class="title">${esc(g.label)}</span><br><span class="meta">${esc(detailOf(g) || '沒有寫尺寸')}・目前 ${g.boxes} 個</span></span>
                ${g === mine ? '<span class="muted">✓ 目前</span>' : ''}
              </button>`,
            )
            .join('')}
          <button class="row" id="q-new"><span class="swatch" style="--c:#8e8e93;pointer-events:none">＋</span><span class="grow"><span class="title">新的一種…</span><br><span class="meta">例如同樣是三通，但尺寸不一樣</span></span></button>
        </div>
        <div id="q-form" style="display:${formOpen ? 'block' : 'none'}"></div>
        <div class="row-actions" style="margin-top:14px">
          <button class="btn secondary" id="q-prev" style="flex:1" ${cur === 0 ? 'disabled' : ''}>‹ 上一個</button>
          <button class="btn secondary" id="q-next" style="flex:1">${cur + 1 >= order.length ? '完成' : '下一個 ›'}</button>
        </div>
        <button class="btn plain block" id="q-more" style="margin-top:8px">改品牌、型號，或存成樣品照…</button>
        <button class="btn danger block" id="q-del" style="margin-top:8px">這不是商品，刪掉這個框</button>`
      body.querySelector('#q-same')?.addEventListener('click', async () => {
        o.checked = true
        await save()
        render()
        go(cur + 1)
      })
      cropBox(photo, o.box)
        .then((blob) => {
          if (previewUrl) URL.revokeObjectURL(previewUrl)
          previewUrl = URL.createObjectURL(blob)
          const img = body.querySelector('#q-img')
          if (img) img.src = previewUrl
        })
        .catch(() => {})
      body.querySelectorAll('[data-q]').forEach((b) =>
        b.addEventListener('click', async () => {
          const g = groups[Number(b.dataset.q)]
          if (g === mine) {
            // 選了目前這一種＝確認一樣
            o.checked = true
            await save()
            return go(cur + 1)
          }
          assign({ label: g.label, brand: g.brand, model: g.model, spec: g.spec })
        }),
      )
      body.querySelector('#q-new').onclick = async () => {
        formOpen = true
        const form = body.querySelector('#q-form')
        form.style.display = 'block'
        form.innerHTML = `<p class="sheet-sub" style="margin:12px 2px 0">品名可以一樣，<b>尺寸要填不一樣的</b>（例如「4分 等徑」「5分×3分 異徑」），才會變成新的一種。</p>${fieldsHtml({ label: o.label, brand: o.brand, model: o.model }, await suggestions())}<button class="btn block" id="q-create" style="margin-top:12px">建立並改成這一種</button>`
        form.scrollIntoView({ block: 'start', behavior: 'smooth' })
        form.querySelector('#f-spec').focus()
        form.querySelector('#q-create').onclick = () => {
          const v = readFields(form)
          if (!v.label) return toast('品名不能空白')
          // 填的跟某一種完全一樣：不是新的一種
          const same = groups.find((g) => g.key === keyOf(v))
          if (same === mine) return toast('跟目前這一種一樣：請填不同的尺寸（例如 4分 等徑）')
          if (same) toast(`已經有這一種（第 ${groups.indexOf(same) + 1} 種），直接改成它`)
          assign(v)
        }
      }
      body.querySelector('#q-prev').onclick = () => go(cur - 1)
      body.querySelector('#q-next').onclick = () => go(cur + 1)
      body.querySelector('#q-del').onclick = async () => {
        const oi = photo.objects.indexOf(o)
        if (oi >= 0) removeObjects(s, [{ pi, oi }])
        order.splice(cur, 1)
        await save()
        render()
        toast('已刪掉這個框')
        if (!order.length) return finish()
        go(Math.min(cur, order.length - 1))
      }
      body.querySelector('#q-more').onclick = () => {
        const oi = photo.objects.indexOf(o)
        finish()
        if (oi >= 0) objectSheet(pi, oi)
      }
    }
    draw()
  }, () => {
    // 點旁邊暗處關掉時，也把標示拿掉
    state.focusObj = null
    if (previewUrl) URL.revokeObjectURL(previewUrl)
    render()
  })
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

// ───────────────────────── 品項庫的面板 ─────────────────────────
/** 搜尋框：只把不符合的列藏起來（不重畫畫面，打字、選字不會被打斷） */
function filterRows(root, q) {
  const cq = canon(q)
  let shown = 0
  root.querySelectorAll('[data-search]').forEach((r) => {
    const ok = !cq || r.dataset.search.includes(cq)
    r.hidden = !ok
    if (ok) shown++
  })
  root.querySelectorAll('.item-sec').forEach((sec) => (sec.hidden = ![...sec.querySelectorAll('[data-search]')].some((r) => !r.hidden)))
  const empty = root.querySelector('#search-empty')
  if (empty) empty.hidden = shown > 0 || !cq
}

/** 選一個品項（合併、設定可以互換） */
function pickItem(title, sub, exclude, onPick) {
  const items = (itemsCache || []).filter((x) => !exclude.includes(x))
  sheet(
    `<h2 class="sheet-title">${esc(title)}</h2><p class="sheet-sub">${esc(sub)}</p>
     <input class="field search" id="pick-q" type="search" placeholder="搜尋品名、型號、料號" autocomplete="off">
     <div class="group" style="margin-top:10px">${
       items.map((it) => `<button class="row" data-pick="${it.id}" data-search="${esc(canon([it.no, it.label, it.brand, it.model, it.spec].join(' ')))}">${itemThumb(it, 36)}<span class="grow"><span class="title">${esc(itemTitle(it))}</span><br><span class="meta">${esc([it.no, it.brand, it.model].filter(Boolean).join('・'))}</span></span></button>`).join('') ||
       '<div class="row muted">品項庫裡沒有其他品項</div>'
     }</div>`,
    (el, close) => {
      el.querySelector('#pick-q').addEventListener('input', (e) => filterRows(el, e.target.value))
      el.querySelectorAll('[data-pick]').forEach((b) =>
        b.addEventListener('click', async () => {
          close()
          await onPick(items.find((x) => x.id === b.dataset.pick))
        }),
      )
    },
  )
}

/** 輸入一個數字（進貨、賣出、帳面數）；quick＝一鍵帶入的數字 */
function numberSheet({ title, sub, value = '', action = '儲存', quick = [] }, onSave) {
  sheet(
    `<h2 class="sheet-title">${esc(title)}</h2><p class="sheet-sub">${sub}</p>
     <input class="field big-num" id="n-val" inputmode="numeric" value="${esc(value)}" aria-label="${esc(title)}">
     ${quick.length ? `<div class="chips" style="margin-top:10px">${quick.map((q) => `<button class="chip" data-n="${q.n}">${esc(q.label)}</button>`).join('')}</div>` : ''}
     <button class="btn block" id="n-save" style="margin-top:14px">${esc(action)}</button>`,
    (el, close) => {
      const input = el.querySelector('#n-val')
      input.focus()
      input.select()
      const done = async (v) => {
        const n = parseInt(v, 10)
        if (!Number.isFinite(n) || n < 0) return toast('請輸入 0 以上的數字')
        close()
        await onSave(n)
        render()
      }
      el.querySelector('#n-save').onclick = () => done(input.value)
      input.addEventListener('keydown', (e) => e.key === 'Enter' && done(input.value))
      el.querySelectorAll('[data-n]').forEach((b) => (b.onclick = () => done(b.dataset.n)))
    },
  )
}
const currentItem = async () => (await itemsAll()).find((x) => x.id === state.itemId)

/** 新增／編輯品項：品名、品牌、型號、規格 */
async function itemEditSheet(it) {
  const sug = await suggestions()
  sheet(
    `<h2 class="sheet-title">${it ? '編輯品項' : '新增品項'}</h2>
     <p class="sheet-sub">${it ? `${esc(it.no)}：改了名稱，以前的寫法還是會對到這一項。` : '型號打對，就會自動解讀、找替代品。'}</p>
     ${fieldsHtml(it || {}, sug)}
     <button class="btn block" id="i-save" style="margin-top:16px">${it ? '儲存' : '新增'}</button>`,
    (el, close) => {
      if (!it) el.querySelector('#f-label').focus()
      el.querySelector('#i-save').onclick = async () => {
        const v = readFields(el)
        if (!v.label) return toast('品名不能空白')
        v.model = normalizeModel(v.model)
        const items = await itemsAll()
        const dup = findItem(items, v)
        if (dup && dup !== it) {
          if (!confirm(`品項庫已經有「${itemTitle(dup)}」（${dup.no}）。${it ? '要把這一項合併過去嗎？' : '要打開它嗎？'}`)) return
          close()
          if (it) {
            await mergeItems(it, dup)
            toast(`已合併到 ${dup.no}`)
          }
          state.itemId = dup.id
          return go('item')
        }
        let target = it
        if (it) {
          Object.assign(it, v, { status: 'ok' })
          if (!it.aliases.includes(looseKey(v))) it.aliases.push(looseKey(v))
        } else target = newItem(items, v, { status: 'ok' })
        await putItem(target)
        close()
        state.itemId = target.id
        go('item')
        toast(it ? '已儲存' : `已新增 ${target.no}`)
      }
    },
  )
}

/** 品項庫右上「⋯」：匯入、匯出、叫貨清單、備份 */
function itemsMoreSheet() {
  sheet(
    `<h2 class="sheet-title">品項庫</h2>
     <div class="group">
       <button class="row edit-only" id="m-import"><span class="grow"><span class="title">貼上 Excel 清單</span><br><span class="meta">一次匯入品名、型號、帳面數、剩幾個要叫貨</span></span>${chev}</button>
       <button class="row" id="m-xlsx"><span class="grow"><span class="title">下載品項庫 Excel</span><br><span class="meta">實盤、帳面、差異、叫貨清單</span></span>${chev}</button>
       <button class="row" id="m-order"><span class="grow"><span class="title">複製叫貨清單</span><br><span class="meta">貼到 LINE 給廠商或老闆</span></span>${chev}</button>
     </div>
     <p class="section-title">備份（換手機、手機壞掉時用）</p>
     <div class="group">
       <button class="row" id="m-backup"><span class="grow"><span class="title">備份品項庫</span><br><span class="meta">下載一個檔案：品項、儲位、樣品照（不含盤點照片）</span></span>${chev}</button>
       <label class="row edit-only"><span class="grow"><span class="title">從備份還原</span><br><span class="meta">選之前下載的備份檔，跟現有的合併</span></span>${chev}<input type="file" accept="application/json,.json" id="m-restore" class="sr-only"></label>
     </div>
     <p class="footnote">品項庫只存在這支手機；建議每週備份一次，或連結 Google 試算表自動同步。</p>`,
    (el, close) => {
      el.querySelector('#m-import').onclick = () => {
        close()
        importSheet()
      }
      el.querySelector('#m-xlsx').onclick = async () => {
        close()
        downloadBlob(itemsXlsx(await itemsAll(true), await whoOfStock()), `品項庫_${ymd(Date.now())}.xlsx`)
        toast('已下載 Excel')
      }
      el.querySelector('#m-order').onclick = async () => {
        close()
        const list = (await itemsAll()).filter(needsOrder)
        if (!list.length) return toast('沒有該叫貨的（要先在品項裡設「剩幾個就要叫貨」）')
        try {
          await navigator.clipboard.writeText(orderText(list))
          toast(`已複製 ${list.length} 項叫貨清單`)
        } catch {
          toast('這個瀏覽器不讓複製，請改用下載 Excel')
        }
      }
      el.querySelector('#m-backup').onclick = async () => {
        close()
        downloadBlob(await backupBlob(), `拍照盤點備份_${ymd(Date.now())}.json`)
        toast('已下載備份檔：存到雲端硬碟或傳給自己')
      }
      el.querySelector('#m-restore').addEventListener('change', async (e) => {
        const f = e.target.files[0]
        if (!f) return
        close()
        await restoreBackup(f)
      })
    },
  )
}
const orderText = (list) =>
  [`叫貨清單 ${ymd(Date.now())}`, ...list.map((it) => `・${itemTitle(it)}${it.brand || it.model ? `（${[it.brand, it.model].filter(Boolean).join(' ')}）` : ''}：剩 ${expected(it)} 個（設定剩 ${it.safety} 個以下要叫貨）`)].join('\n')

/** 品項庫的表格（Excel、Google 試算表共用） */
const ITEM_HEAD = ['料號', '品名', '品牌', '型號', '尺寸／規格', '實盤', '帳面', '差異', '剩幾個要叫貨（安全庫存）', '狀態', '在哪裡（位置 數量）', '最近盤點', '盤點人']
/** 每一格「誰盤的」：從那次盤點找盤點人 */
async function whoOfStock() {
  const byId = new Map((await db.all()).map((s) => [s.id, s]))
  return (st) => (byId.get(st.sid) ? byName(byId.get(st.sid)) : '')
}
const timeText = (ms) => (ms ? `${ymd(ms)} ${hm(ms)}` : '')
/** time：時間怎麼寫（Excel 用文字；Google 試算表送 {$t} 讓那邊轉成真的日期） */
function itemRows(items, whoOf = () => '', time = timeText) {
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  return [...items]
    .sort((a, b) => cmp(a.label, b.label) || cmp(a.spec, b.spec) || cmp(a.no, b.no))
    .map((it) => {
      const d = diffOf(it)
      const latest = liveStock(it)
        .map(([, st]) => st)
        .sort((a, b) => b.at - a.at)[0]
      const last = latest?.at
      return [
        it.no,
        it.label,
        it.brand,
        it.model,
        it.spec,
        onHand(it),
        it.book ?? '',
        d ?? '',
        it.safety ?? '',
        [it.status === 'new' ? '新的（待確認）' : '', needsOrder(it) ? '該叫貨' : '', d > 0 ? '盤盈' : d < 0 ? '盤虧' : ''].filter(Boolean).join('、'),
        liveStock(it)
          .map(([, st]) => `${st.place || '沒填位置'} ${st.count}`)
          .join('、'),
        last ? time(last) : '',
        latest ? whoOf(latest) : '',
      ]
    })
}
/** 儲位庫存：每一格放了什麼、幾個、誰什麼時候盤的 */
function stockRows(items, whoOf = () => '', time = timeText) {
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  return items
    .flatMap((it) => liveStock(it).map(([, st]) => [st.place || '（沒填位置）', findLocation(st.place)?.name || '', it.no, it.label, it.brand, it.model, it.spec, st.count, time(st.at), whoOf(st)]))
    .sort((a, b) => cmp(a[0], b[0]) || cmp(a[3], b[3]) || cmp(a[6], b[6]))
}
const STOCK_HEAD = ['儲位', '儲位說明', '料號', '品名', '品牌', '型號', '尺寸／規格', '數量', '盤點時間', '盤點人']
/** 儲位盤點進度：每一格最近什麼時候盤、誰盤的；還沒盤、超過 30 天沒盤的標出來 */
function locationRows(items, whoOf = () => '', time = timeText) {
  const seen = new Map()
  for (const it of items)
    for (const [, st] of liveStock(it)) {
      const k = canon(st.place || '（沒填位置）')
      const cur = seen.get(k) || { place: st.place || '（沒填位置）', items: 0, qty: 0, at: 0, st: null }
      cur.items += 1
      cur.qty += st.count
      if (st.at > cur.at) Object.assign(cur, { at: st.at, st })
      seen.set(k, cur)
    }
  const list = locations().map((l) => ({ code: l.code, name: l.name || '', info: seen.get(canon(l.code)) }))
  for (const [k, info] of seen) if (!locations().some((l) => canon(l.code) === k)) list.push({ code: info.place, name: '', info })
  const month = 30 * 24 * 3600 * 1000
  return list.map(({ code, name, info }) => [code, name, info?.items || 0, info?.qty || 0, info ? time(info.at) : '', info?.st ? whoOf(info.st) : '', !info ? '還沒盤' : Date.now() - info.at > month ? '超過 30 天沒盤' : '已盤'])
}
const LOC_HEAD = ['儲位', '說明', '品項數', '件數', '最近盤點', '盤點人', '狀態']
const itemSheets = (items, whoOf) => [
  { name: '品項庫', rows: [ITEM_HEAD, ...itemRows(items, whoOf)] },
  { name: '儲位庫存', rows: [STOCK_HEAD, ...stockRows(items, whoOf)] },
  { name: '叫貨清單', rows: [['料號', '品名', '品牌', '型號', '尺寸／規格', '現在大概有', '剩幾個要叫貨（安全庫存）'], ...items.filter(needsOrder).map((it) => [it.no, it.label, it.brand, it.model, it.spec, expected(it), it.safety])] },
]
const itemsXlsx = (items, whoOf) => makeXlsx(itemSheets(items, whoOf))
/** 送給 Google 試算表的報表資料（時間送 {$t}） */
async function sheetReport() {
  const items = await itemsAll(true)
  const whoOf = await whoOfStock()
  const t = (ms) => (ms ? { $t: ms } : '')
  return { items: [ITEM_HEAD, ...itemRows(items, whoOf, t)], stock: [STOCK_HEAD, ...stockRows(items, whoOf, t)], locs: [LOC_HEAD, ...locationRows(items, whoOf, t)], at: Date.now() }
}

/** 貼上 Excel 清單：第一列是標題，看標題認欄位 */
function parseTable(text) {
  const lines = String(text).split(/\r?\n/).filter((l) => l.trim())
  if (lines.length < 2) throw new Error('至少要有標題列和一列資料')
  const sep = lines[0].includes('\t') ? '\t' : ','
  const rows = lines.map((l) => l.split(sep).map((c) => c.trim().replace(/^"(.*)"$/, '$1')))
  const head = rows[0].map((h) => h.replace(/\s/g, ''))
  const col = (...ws) => head.findIndex((h) => ws.some((w) => h.includes(w)))
  const safety = col('安全', '最低', '叫貨')
  const bookWords = ['帳面', '庫存', '數量', '存量']
  const c = { no: col('料號', '編號'), label: col('品名', '名稱', '品項'), brand: col('品牌', '廠牌'), model: col('型號'), spec: col('規格', '尺寸'), book: head.findIndex((h, i) => i !== safety && bookWords.some((w) => h.includes(w))), safety }
  if (c.label < 0) throw new Error('第一列要有標題，而且要有「品名」')
  const num = (v) => {
    const n = parseInt(String(v ?? '').replace(/[^\d-]/g, ''), 10)
    return Number.isFinite(n) ? n : null
  }
  const at = (r, i) => (i >= 0 ? r[i] || '' : '')
  return rows.slice(1).map((r) => ({ no: at(r, c.no), label: at(r, c.label), brand: at(r, c.brand), model: normalizeModel(at(r, c.model)), spec: at(r, c.spec), book: c.book >= 0 ? num(r[c.book]) : null, safety: c.safety >= 0 ? num(r[c.safety]) : null }))
}
async function importTable(text) {
  const rows = parseTable(text)
  const items = await itemsAll()
  let created = 0
  let updated = 0
  let skipped = 0
  for (const r of rows) {
    if (!r.label) {
      skipped++
      continue
    }
    let it = (r.no && items.find((x) => canon(x.no) === canon(r.no))) || findItem(items, r)
    if (!it) {
      it = newItem(items, r, { status: 'ok', ...(r.no ? { no: r.no } : {}) })
      items.unshift(it)
      created++
    } else {
      updated++
      it.status = 'ok'
      if (!it.aliases.includes(looseKey(r))) it.aliases.push(looseKey(r))
    }
    if (r.book != null && r.book !== it.book) {
      it.book = r.book
      it.moves.push({ at: Date.now(), kind: 'set', qty: r.book })
    }
    if (r.safety != null) it.safety = r.safety
    await putItem(it)
  }
  return { created, updated, skipped }
}
function importSheet() {
  sheet(
    `<h2 class="sheet-title">貼上 Excel 清單</h2>
     <p class="sheet-sub">在 Excel 或 Google 試算表選取整個表格（含第一列標題）→ 複製 → 貼在下面。標題要有「品名」；有這些欄位也會讀：料號、品牌、型號、規格（或尺寸）、帳面（或庫存、數量）、安全庫存。</p>
     <textarea class="field" id="imp-text" placeholder="品名&#9;型號&#9;規格&#9;帳面數&#9;安全庫存&#10;乾燥過濾器&#9;DML 083S&#9;3分&#9;12&#9;5"></textarea>
     <button class="btn block" id="imp-go" style="margin-top:12px">匯入</button>
     <p class="footnote">同名稱、同型號的會更新，不會重複；料號欄有寫就用你的料號。</p>`,
    (el, close) => {
      el.querySelector('#imp-go').onclick = async () => {
        try {
          const r = await importTable(el.querySelector('#imp-text').value)
          close()
          toast(`匯入完成：新增 ${r.created}、更新 ${r.updated}${r.skipped ? `、略過 ${r.skipped} 列（沒有品名）` : ''}`)
          go('items')
        } catch (e) {
          toast(e.message)
        }
      }
    },
  )
}

/** 備份：品項（含小圖）、儲位、品項清單、樣品照，存成一個 JSON 檔 */
const b64ToBlob = (b64, type = 'image/jpeg') => {
  const bin = atob(b64)
  const arr = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
  return new Blob([arr], { type })
}
async function backupBlob() {
  const items = await itemsAll(true)
  const samples = await idb.samples.all().catch(() => [])
  const enc = async (b) => (b ? await blobToBase64(b) : undefined)
  const data = {
    app: '拍照盤點',
    kind: 'backup',
    version: VERSION,
    at: Date.now(),
    locations: locations(),
    catalog: ls.get(LS.catalog),
    items: await Promise.all(items.map(async (it) => ({ ...it, photo: await enc(it.photo) }))),
    samples: await Promise.all(samples.map(async (s) => ({ ...s, blob: await enc(s.blob) }))),
  }
  return new Blob([JSON.stringify(data)], { type: 'application/json' })
}
async function restoreBackup(file) {
  let data
  try {
    data = JSON.parse(await file.text())
  } catch {
    return toast('這個檔案打不開')
  }
  if (data?.app !== '拍照盤點' || data.kind !== 'backup') return toast('這不是拍照盤點的備份檔')
  if (!confirm(`還原 ${fmtTime(data.at)} 的備份：${data.items?.length || 0} 個品項、${data.locations?.length || 0} 個儲位、${data.samples?.length || 0} 張樣品照。\n同一個品項以備份為準，其他的保留。`)) return
  for (const it of data.items || []) await idb.items.put({ ...it, photo: it.photo ? b64ToBlob(it.photo) : undefined })
  for (const s of data.samples || []) if (s.blob) await idb.samples.put({ ...s, blob: b64ToBlob(s.blob) })
  const locs = locations()
  for (const l of data.locations || []) if (!locs.some((x) => canon(x.code) === canon(l.code))) locs.push(l)
  saveLocations(locs)
  if (data.catalog && !ls.get(LS.catalog)) ls.set(LS.catalog, data.catalog)
  itemsCache = null
  toast('已還原')
  go('items')
}

/** 儲位：新增（可以一次建立一整排）／修改 */
function locationSheet(index = -1) {
  const list = locations()
  const loc = list[index]
  sheet(
    `<h2 class="sheet-title">${loc ? '修改儲位' : '新增儲位'}</h2>
     <div class="form">
       <label class="field-label" for="l-code">代號（印在標籤上，越短越好）</label>
       <input class="field" id="l-code" value="${esc(loc ? loc.code : nextCode(list))}" autocomplete="off" autocapitalize="characters">
       <label class="field-label" for="l-name">說明（選填）</label>
       <input class="field" id="l-name" value="${esc(loc?.name || '')}" placeholder="例：冷凍油、冷媒那一排第 1 層" autocomplete="off">
     </div>
     <button class="btn block" id="l-save" style="margin-top:14px">${loc ? '儲存' : '新增'}</button>
     ${
       loc
         ? '<button class="btn danger block" id="l-del" style="margin-top:10px">刪除這個儲位</button><p class="footnote">刪除儲位不會刪掉品項的數量。</p>'
         : `<p class="section-title">或一次建立一整排</p>
            <div class="row-actions"><input class="field" id="l-row" value="${esc((/^[A-Za-z]+/.exec(nextCode(list)) || ['A'])[0])}" style="flex:1" aria-label="排的字母" autocapitalize="characters"><input class="field" id="l-n" inputmode="numeric" value="5" style="flex:1" aria-label="幾層"></div>
            <p class="footnote">左邊填排的字母、右邊填幾層：例如 A、5 → A-01～A-05。</p>
            <button class="btn secondary block" id="l-batch" style="margin-top:8px">建立整排</button>`
     }`,
    (el, close) => {
      el.querySelector('#l-save').onclick = () => {
        const code = el.querySelector('#l-code').value.trim().toUpperCase()
        const name = el.querySelector('#l-name').value.trim()
        if (!code) return toast('代號不能空白')
        if (list.some((l, i) => i !== index && canon(l.code) === canon(code))) return toast(`已經有 ${code} 了`)
        if (loc) Object.assign(loc, { code, name })
        else list.push({ code, name })
        saveLocations(list)
        close()
        render()
      }
      el.querySelector('#l-del')?.addEventListener('click', () => {
        if (!confirm(`刪除儲位 ${loc.code}？`)) return
        list.splice(index, 1)
        saveLocations(list)
        close()
        render()
      })
      el.querySelector('#l-batch')?.addEventListener('click', () => {
        const row = el.querySelector('#l-row').value.trim().toUpperCase()
        const n = Math.min(30, parseInt(el.querySelector('#l-n').value, 10) || 0)
        if (!row || !n) return toast('請填排的字母和層數')
        let added = 0
        for (let i = 1; i <= n; i++) {
          const code = `${row}-${String(i).padStart(2, '0')}`
          if (list.some((l) => canon(l.code) === canon(code))) continue
          list.push({ code, name: `${row} 排第 ${i} 層` })
          added++
        }
        saveLocations(list)
        close()
        render()
        toast(added ? `建立了 ${added} 個儲位` : '這一排已經都有了')
      })
    },
  )
}
/** 列印標籤：大字代號（AI 拍到就認得），A4 一頁 10 張 */
function printLabels() {
  document.getElementById('print-area')?.remove()
  const area = document.createElement('div')
  area.id = 'print-area'
  area.innerHTML = locations()
    .map((l) => `<div class="label"><div class="label-code">${esc(l.code)}</div>${l.name ? `<div class="label-name">${esc(l.name)}</div>` : ''}<div class="label-foot">拍照盤點・儲位</div></div>`)
    .join('')
  document.body.appendChild(area)
  window.print()
  setTimeout(() => area.remove(), 1000)
}

// ───────────────────────── 查型號：拍標籤給 AI 讀 ─────────────────────────
const LABEL_SCHEMA = {
  type: 'OBJECT',
  properties: {
    label: { type: 'STRING', description: '中文品名，優先用店內品項清單的寫法；看不出來就空字串' },
    brand: { type: 'STRING', description: '品牌；看不出來就空字串' },
    model: { type: 'STRING', description: '型號，照標籤原樣（例如 DML 083S、TES 2、KP 15）；看不出來就空字串' },
    spec: { type: 'STRING', description: '尺寸／規格（例如 3分、R404A、220V）；看不出來就空字串' },
    code: { type: 'STRING', description: '訂購碼／料號（例如 Danfoss 023Z5040）；沒有就空字串' },
    text: { type: 'STRING', description: '照片上看得到的字，照原樣抄（最多 300 字）' },
    item_no: { type: 'STRING', description: '就是「品項庫」裡的某一項時填料號；不確定就空字串' },
  },
  required: ['label', 'brand', 'model', 'spec', 'code', 'text'],
}
/** 問 AI 一次、回 JSON：同一個模型先有思考設定、再拿掉；不行就換下一個模型 */
async function askJson(parts, schema, timeoutMs = 30000) {
  const first = await currentModel()
  const list = modelCache ?? (await fetchModels().catch(() => [first]))
  const models = [first, ...list.filter((m) => m !== first)].slice(0, 3)
  let lastErr
  for (const model of models) {
    for (const think of [true, false]) {
      try {
        const generationConfig = { temperature: 0, responseMimeType: 'application/json', responseSchema: schema, ...(think && thinkingFor(model) ? { thinkingConfig: thinkingFor(model) } : {}) }
        const data = await call(`models/${model}:generateContent`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contents: [{ role: 'user', parts }], generationConfig }) }, timeoutMs)
        const text = data?.candidates?.[0]?.content?.parts?.map((p) => p.text || '').join('') || ''
        // aiModel：用哪個 AI 模型（不能叫 model，會蓋掉零件的型號欄位）
        return { ...JSON.parse(text.replace(/^```(?:json)?\s*|\s*```$/g, '')), aiModel: model }
      } catch (e) {
        lastErr = e
        if (e instanceof SyntaxError) break
        if (!(e instanceof ApiError)) throw e
        if (e.status === 400 && think) continue
        if (e.status === 400 || !(retryable(e) || e.status === 404)) throw e
        break
      }
    }
  }
  throw lastErr instanceof SyntaxError ? new ApiError('AI 回傳的格式壞掉了，請再拍一次。', 0) : lastErr
}
async function readLabel(file) {
  const photo = await prepareImage(file)
  const thumb = await sampleFromFile(file)
  const prompt = `你是冷凍空調材料行的店員。照片是一個零件的標籤、銘牌、外盒，或零件本體上的刻字（可能是客人帶來的舊零件）。
請讀出：label＝中文品名、brand＝品牌、model＝型號（照標籤原樣）、spec＝尺寸／規格、code＝訂購碼、text＝看得到的字全部照抄。看不清楚就空字串，不要猜。
店內品項清單：
${catalogLines().join('\n')}${itemsPrompt()}`
  const r = await askJson([{ inline_data: { mime_type: 'image/jpeg', data: await blobToBase64(photo.blob) } }, { text: prompt }], LABEL_SCHEMA)
  const it = promptItems().find((x) => x.no.toLowerCase() === String(r.item_no || '').trim().toLowerCase())
  const read = it
    ? { label: it.label, brand: it.brand, model: it.model, spec: it.spec }
    : { label: cleanLabel(r.label || ''), brand: String(r.brand || '').trim(), model: normalizeModel(r.model), spec: String(r.spec || '').trim() }
  // AI 沒寫品名、但型號看得懂（例如 DML 083S）→ 用解讀出來的品名、品牌、規格補上
  const d = decodedFor({ ...read, code: r.code, text: r.text }).find((x) => x.label)
  if (d) for (const f of ['label', 'brand', 'model', 'spec']) if (!read[f] && d[f]) read[f] = d[f]
  return { ...read, code: String(r.code || '').trim(), text: String(r.text || '').trim().slice(0, 300), thumb, url: URL.createObjectURL(thumb) }
}

// ───────────────────────── 多台裝置同步（存在使用者自己的 Google 雲端硬碟） ─────────────────────────
/**
 * 手機拍完、電腦打開也看得到：每一筆（一次盤點含照片、品項、樣品照、儲位設定）各自同步。
 * - 上傳：本機改過、還沒同步的（updatedAt ≠ _syncT）
 * - 下載：別台送上去、比上次新的；兩邊都改過就以「比較晚改的」為準
 * - 刪除也會同步（墓碑：記下刪了哪一筆、什麼時候）
 * 要有 Apps Script 網址＋同步密碼才讀得到，資料不會公開。
 */
const deviceId = () => {
  let id = ls.get(LS.device)
  if (!id) {
    id = uid()
    ls.set(LS.device, id)
  }
  return id
}
const readJson = (k, d) => {
  try {
    return JSON.parse(ls.get(k, '')) ?? d
  } catch {
    return d
  }
}
function tombstone(k) {
  const list = readJson(LS.deleted, [])
  list.push({ k, t: Date.now() })
  ls.set(LS.deleted, JSON.stringify(list.slice(-2000)))
  scheduleSync()
}
const syncReady = () => !!(ls.get(LS.sheet) && ls.get(LS.syncKey))
let syncTimer = 0
let syncing = null
/** 改了東西：等一下（合併連續的修改）再同步；辨識中先不要 */
function scheduleSync(ms = 6000) {
  if (!syncReady()) return
  clearTimeout(syncTimer)
  syncTimer = setTimeout(() => {
    if (state.moving) return
    if (state.view === 'analyzing' || syncing) return scheduleSync(ms)
    syncNow().catch(() => {})
  }, ms)
}
/** 大家自動同步：App 開著時每 40 秒看一次有沒有別人的新資料；切回 App 時馬上看一次 */
setInterval(() => {
  if (syncReady() && document.visibilityState === 'visible' && !syncing && !state.moving && state.view !== 'analyzing') syncNow().catch(() => {})
}, 40000)
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState !== 'visible' || !syncReady() || syncing || state.moving || state.view === 'analyzing') return
  if (Date.now() - Number(ls.get(LS.lastSync, '0')) > 10000) syncNow().catch(() => {})
})
async function postSync(body) {
  let res
  try {
    res = await fetch(ls.get(LS.sheet), { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ ...body, key: ls.get(LS.syncKey), dev: deviceId() }) })
  } catch {
    throw new Error('連不到 Google（沒有網路？）')
  }
  let data
  try {
    data = await res.json()
  } catch {
    throw new Error('Google 試算表沒有正確回覆：請確認貼的是新版程式碼，而且部署成「新版本」')
  }
  // 擁有者把資料搬到新的 Google 帳號：自動改用新網址（連結碼不用重貼）
  if (data.moved && !body.followed) {
    ls.set(LS.sheet, data.moved)
    toast('資料搬到新的位置了，已自動跟過去')
    return postSync({ ...body, followed: true })
  }
  // 被移除權限：這台的公司資料自動清掉
  if (data.revoked === true && !data.ok) {
    await wipeLocal('這台已經被移除權限，App 裡的資料已清除')
    throw new Error('這台已經被移除權限，App 裡的資料已清除')
  }
  // 權限隨時可能被管理員改：每次回覆都更新
  if (data.me?.role) {
    const changed = data.me.role !== ls.get(LS.memberRole)
    ls.set(LS.memberRole, data.me.role)
    ls.set(LS.memberName, data.me.name || '')
    ls.set(LS.memberId, data.me.id || '')
    if (changed) {
      toast(`你的權限：${ROLE_LABEL[data.me.role]}（${ROLE_DESC[data.me.role]}）`)
      setTimeout(render, 0)
    }
  }
  if (Array.isArray(data.roster) && data.roster.length) {
    const changed = JSON.stringify(data.roster) !== ls.get(LS.roster)
    ls.set(LS.roster, JSON.stringify(data.roster))
    // 選的盤點人被移除了：改回這台自己
    if (ls.get(LS.counterId) && !data.roster.some((p) => p.id === ls.get(LS.counterId))) ls.set(LS.counterId, '')
    if (changed && ['home', 'review', 'settings', 'capture'].includes(state.view)) setTimeout(render, 0)
  }
  if (!data.ok) throw new Error(data.error || '同步失敗')
  return data
}
/** 清除這台的資料（退出同步、被移除權限時） */
async function wipeLocal(msg) {
  clearTimeout(syncTimer)
  for (const k of [LS.syncKey, LS.sheet, LS.pulled, LS.lastSync, LS.deleted, LS.settingsAt, LS.settingsSyncT, LS.locations, LS.key, LS.catalog, LS.memberName, LS.memberRole, LS.memberId, LS.roster, LS.counterId]) ls.set(k, '')
  // 直接清本機（不留刪除紀錄，才不會把雲端的資料也刪掉）
  await idb.sessions.clear()
  await idb.items.clear()
  await idb.samples.clear()
  itemsCache = null
  state.session = null
  toast(msg)
  go('home')
}
const SYNC_STORES = { session: idb.sessions, item: idb.items, sample: idb.samples }
const tOf = (v) => v.updatedAt || v.createdAt || 0
async function encodeRecord(kind, v) {
  if (kind === 'session') return { ...v, photos: await Promise.all(v.photos.map(async (p) => ({ ...p, blob: undefined, b64: await blobToBase64(p.blob) }))) }
  if (kind === 'item') return { ...v, photo: undefined, photoB64: v.photo ? await blobToBase64(v.photo) : '' }
  return { ...v, blob: undefined, b64: await blobToBase64(v.blob) }
}
function decodeRecord(kind, d) {
  if (kind === 'session') return { ...d, photos: (d.photos || []).map(({ b64, ...p }) => ({ ...p, blob: b64ToBlob(b64 || '') })) }
  if (kind === 'item') {
    const { photoB64, ...rest } = d
    return { ...rest, photo: photoB64 ? b64ToBlob(photoB64) : undefined }
  }
  const { b64, ...rest } = d
  return { ...rest, blob: b64ToBlob(b64 || '') }
}
/** 合併同一個品項的兩個版本：名稱、帳面設定等以比較晚改的為準；各位置數量逐格比；別名、進出紀錄、可互換合起來 */
function mergeItemData(local, remote) {
  const base = tOf(remote) >= tOf(local) ? remote : local
  const stock = {}
  for (const k of new Set([...Object.keys(local.stock || {}), ...Object.keys(remote.stock || {})])) stock[k] = newerEntry(local.stock?.[k], remote.stock?.[k])
  const moves = [...new Map([...(local.moves || []), ...(remote.moves || [])].map((m) => [`${m.at}|${m.kind}|${m.qty}`, m])).values()].sort((a, b) => a.at - b.at)
  const book = moves.length ? bookFromMoves(moves) : base.book
  return {
    ...base,
    stock,
    moves,
    book,
    aliases: [...new Set([...(local.aliases || []), ...(remote.aliases || [])])],
    equiv: [...new Set([...(local.equiv || []), ...(remote.equiv || [])])],
    photo: base.photo || local.photo || remote.photo,
  }
}
/** 比較兩個版本內容是否一樣（不看照片、時間） */
const itemSig = (it) =>
  JSON.stringify([
    it.no,
    it.label,
    it.brand,
    it.model,
    it.spec,
    it.status,
    it.book,
    it.safety,
    Object.entries(it.stock || {}).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)),
    [...(it.aliases || [])].sort(),
    (it.moves || []).map((m) => `${m.at}|${m.kind}|${m.qty}`),
    [...(it.equiv || [])].sort(),
  ])
/** 套用一筆別台的資料；回傳本機有沒有變 */
async function applyRemote(rec) {
  if (rec.k === 'settings') {
    if (!rec.d || rec.t <= Number(ls.get(LS.settingsAt, '0'))) return false
    ls.set(LS.locations, JSON.stringify(rec.d.locations || []))
    if (rec.d.catalog != null) ls.set(LS.catalog, rec.d.catalog)
    ls.set(LS.settingsAt, String(rec.t))
    ls.set(LS.settingsSyncT, String(rec.t))
    return true
  }
  const i = rec.k.indexOf(':')
  const kind = rec.k.slice(0, i)
  const id = rec.k.slice(i + 1)
  const store = SYNC_STORES[kind]
  if (!store) return false
  const cur = await store.get(id)
  if (rec.del) {
    if (!cur || tOf(cur) > rec.t) return false
    await store.del(id)
    if (kind === 'session' && state.session?.id === id && ['review', 'capture'].includes(state.view)) {
      toast('這次盤點在另一台被刪掉了')
      go('home')
    }
    return true
  }
  if (!rec.d) return false
  // 品項：兩台都改過（例如兩個人同時盤到同一種、在不同格）→ 合併，不要互相蓋掉
  if (kind === 'item' && cur) {
    const remote = { ...decodeRecord('item', rec.d), updatedAt: rec.t }
    const merged = mergeItemData(cur, remote)
    if (itemSig(merged) === itemSig(remote)) {
      // 雲端那份已經包含本機的全部：直接用它（標成已同步）
      await idb.items.putRaw({ ...remote, photo: remote.photo || cur.photo, _syncT: rec.t })
      return itemSig(cur) !== itemSig(remote)
    }
    // 本機有雲端沒有的：存合併結果，時間設得比雲端新、標成還沒同步，等一下傳上去讓大家一致
    await idb.items.putRaw({ ...merged, updatedAt: Math.max(Date.now(), rec.t + 1), _syncT: undefined })
    scheduleSync(2000)
    return itemSig(merged) !== itemSig(cur)
  }
  if (cur && tOf(cur) >= rec.t) return false
  const v = { ...decodeRecord(kind, rec.d), updatedAt: rec.t, _syncT: rec.t }
  // 兩台同時建立新品項、料號撞號：本機還沒傳上去的那個改用新號碼
  if (kind === 'item') {
    const items = await idb.items.all()
    for (const x of items) {
      if (x.id !== v.id && x.no === v.no && x._syncT === undefined) {
        x.no = nextNo([...items, v])
        await idb.items.put(x)
      }
    }
  }
  await store.putRaw(v)
  if (kind === 'session' && state.session?.id === id) {
    state.session = v
    if (state.view === 'review') toast('另一台更新了這次盤點')
  }
  return true
}
/** 同步一次（同時只跑一個）：先上傳、再下載 */
async function syncNow(onProgress = () => {}) {
  if (!syncReady()) throw new Error('還沒開啟多台同步')
  if (syncing) return syncing
  syncing = (async () => {
    state.syncState = 'syncing'
    // 1. 上傳：還沒同步過的（檢視者只下載）
    const jobs = []
    const viewer = !canEdit()
    for (const s of await db.all()) if (tOf(s) !== s._syncT && s.photos.some((p) => p.status === 'done')) jobs.push(['session', s])
    for (const it of await itemsAll(true)) if (tOf(it) !== it._syncT) jobs.push(['item', it])
    for (const sm of await idb.samples.all().catch(() => [])) if (tOf(sm) !== sm._syncT) jobs.push(['sample', sm])
    const settingsAt = Number(ls.get(LS.settingsAt, '0'))
    const settingsDirty = settingsAt && String(settingsAt) !== ls.get(LS.settingsSyncT)
    const deleted = readJson(LS.deleted, [])
    const total = jobs.length + deleted.length + (settingsDirty ? 1 : 0)
    let batch = []
    let marks = []
    let size = 0
    let sent = 0
    const flush = async () => {
      if (!batch.length) return
      const r = await postSync({ action: 'push', records: batch })
      // 雲端已經有比較新的（別台改的）→ 不要標成已同步；下面下載時會拿到新的版本再合併
      const skipped = new Set(r.skipped || [])
      for (const [i, m] of marks.entries()) if (!skipped.has(batch[i].k)) await m()
      sent += batch.length
      onProgress(`上傳 ${sent}／${total}`)
      batch = []
      marks = []
      size = 0
    }
    const add = async (rec, mark) => {
      const len = JSON.stringify(rec).length
      if (size && size + len > 5e6) await flush()
      batch.push(rec)
      marks.push(mark)
      size += len
    }
    if (viewer) {
      jobs.length = 0
      deleted.length = 0
    }
    for (const [kind, v] of jobs) {
      const t = tOf(v)
      await add({ k: `${kind}:${v.id}`, t, d: await encodeRecord(kind, v) }, async () => {
        await SYNC_STORES[kind].markSynced(v.id, t)
        if (kind === 'item') {
          const c = itemsCache?.find((x) => x.id === v.id)
          if (c && tOf(c) === t) c._syncT = t
        }
        if (kind === 'session' && state.session?.id === v.id && tOf(state.session) === t) state.session._syncT = t
      })
    }
    if (settingsDirty && !viewer) await add({ k: 'settings', t: settingsAt, d: { locations: locations(), catalog: ls.get(LS.catalog) } }, async () => ls.set(LS.settingsSyncT, String(settingsAt)))
    for (const d of deleted)
      await add({ k: d.k, t: d.t, del: true }, async () => {
        ls.set(LS.deleted, JSON.stringify(readJson(LS.deleted, []).filter((x) => !(x.k === d.k && x.t === d.t))))
      })
    await flush()
    // 2. 下載：別台送上去、比上次新的
    let cursor = Number(ls.get(LS.pulled, '0'))
    let got = 0
    let changed = false
    for (let round = 0; round < 200; round++) {
      const r = await postSync({ action: 'pull', since: cursor })
      for (const rec of r.records) if (await applyRemote(rec)) changed = true
      got += r.records.length
      if (r.records.length) onProgress(`下載 ${got} 筆`)
      cursor = r.next
      ls.set(LS.pulled, String(cursor))
      if (!r.more) break
    }
    ls.set(LS.lastSync, String(Date.now()))
    state.syncState = ''
    if (changed) {
      itemsCache = null
      if (['home', 'items', 'item', 'report', 'locations', 'review'].includes(state.view)) render()
    } else if (state.view === 'home') render()
    return { pushed: sent, pulled: got, changed }
  })()
    .catch((e) => {
      state.syncState = 'error'
      state.syncError = e.message
      if (state.view === 'home') render()
      throw e
    })
    .finally(() => (syncing = null))
  return syncing
}
/** 連結碼＝網址＋同步密碼（給另一台貼上） */
const linkCode = () => `${ls.get(LS.sheet)}#k=${ls.get(LS.syncKey)}`
function parseLinkCode(text) {
  const m = /^(https:\/\/script\.google\.com\/macros\/s\/[^#\s]+\/exec)#k=([\w-]{12,})$/.exec(String(text).trim())
  return m ? { url: m[1], key: m[2] } : null
}
const newSyncKey = () => {
  const a = new Uint8Array(18)
  crypto.getRandomValues(a)
  return [...a].map((b) => b.toString(36).padStart(2, '0')).join('').slice(0, 24)
}
/** 共用設定（像 Google 雲端硬碟的「共用」視窗）：新增使用者＋權限、名單、一般存取權 */
function shareSheet() {
  const appUrl = `${location.origin}${location.pathname}`
  const personRow = (p) => {
    const me = p.id === (ls.get(LS.memberId) || (myRole() === 'owner' ? 'owner' : ''))
    const meta = p.seen ? `最後上線 ${fmtTime(p.seen)}` : '還沒加入'
    const right =
      p.role === 'owner'
        ? `<span class="muted role-text">擁有者</span>${me ? '<button class="btn small plain" data-rename-me>改名字</button>' : ''}`
        : me
          ? `<span class="muted role-text">${ROLE_LABEL[p.role]}</span>`
          : `<select class="role-select" data-member="${esc(p.id)}" data-name="${esc(p.name)}" data-role="${p.role}" aria-label="${esc(p.name)} 的權限">
              ${['manager', 'editor', 'viewer'].map((r) => `<option value="${r}" ${r === p.role ? 'selected' : ''}>${ROLE_LABEL[r]}</option>`).join('')}
              <option disabled>──────</option>
              <option value="rename">改名字…</option>
              <option value="reissue">重新產生連結碼…</option>
              <option value="remove">移除權限</option>
            </select>`
    return `<div class="row person"><span class="avatar" aria-hidden="true">${esc(p.name.slice(0, 1))}</span><span class="grow"><span class="title">${esc(p.name)}${me ? '（你）' : ''}</span><br><span class="meta">${meta}</span></span>${right}</div>`
  }
  sheet(
    `<h2 class="sheet-title">共用「拍照盤點」</h2>
     <div class="form">
       <input class="field" id="sh-name" placeholder="新增使用者：名字（例如 阿明、辦公室電腦）" autocomplete="off">
       <div class="row-actions"><select class="field role-pick" id="sh-role" aria-label="權限"><option value="editor">編輯者（可以盤點、修改）</option><option value="viewer">檢視者（只能看）</option><option value="manager">管理員（也可以加人）</option></select><button class="btn" id="sh-invite">邀請</button></div>
     </div>
     <p class="section-title">擁有存取權的使用者</p>
     <div class="group" id="sh-list"><div class="row muted">載入中…</div></div>
     <p class="section-title">一般存取權</p>
     <div class="group"><div class="row"><span class="avatar lock" aria-hidden="true">🔒</span><span class="grow"><span class="title">限制</span><br><span class="meta">只有上面名單裡的人，用自己的連結碼才能開啟</span></span></div></div>
     <details class="steps"><summary>權限說明</summary><ol>${Object.keys(ROLE_LABEL)
       .map((r) => `<li><b>${ROLE_LABEL[r]}：</b>${ROLE_DESC[r]}</li>`)
       .join('')}<li>有人離職：選「移除權限」，只有他失效；他的手機下次連線，App 裡的資料會自動清除。其他人不用改。</li><li>換手機或連結碼外流：選「重新產生連結碼」，舊的馬上失效。</li></ol></details>
     ${myRole() === 'owner' ? '<button class="btn plain block" id="sh-move" style="margin-top:14px">搬到另一個 Google 帳號（例如公司帳號）</button>' : ''}`,
    (el, close) => {
      const list = el.querySelector('#sh-list')
      const load = async () => {
        let people
        try {
          people = (await postSync({ action: 'members' })).people
        } catch (e) {
          list.innerHTML = `<div class="row muted">${esc(e.message)}</div>`
          return
        }
        list.innerHTML = people.map(personRow).join('')
        list.querySelector('[data-rename-me]')?.addEventListener('click', () => {
          close()
          nameSheet('owner', '你的名字？', '大家會在盤點紀錄看到是誰盤的。')
        })
        list.querySelectorAll('select[data-member]').forEach((sel) =>
          sel.addEventListener('change', async () => {
            const { member: id, name, role } = sel.dataset
            const v = sel.value
            sel.value = role // 先還原，成功才更新
            try {
              if (['manager', 'editor', 'viewer'].includes(v)) {
                await postSync({ action: 'setRole', id, role: v })
                toast(`「${name}」改成${ROLE_LABEL[v]}，下次同步就生效`)
              } else if (v === 'rename') {
                close()
                return nameSheet(id, `改「${name}」的名字`, '')
              } else if (v === 'reissue') {
                if (!confirm(`幫「${name}」重新產生連結碼？\n舊的連結碼馬上失效（他舊手機上的資料會在下次連線時清除）。`)) return
                const key = newSyncKey()
                await postSync({ action: 'reissue', id, newKey: key })
                close()
                return linkSheet(name, role, `${ls.get(LS.sheet)}#k=${key}`, appUrl)
              } else if (v === 'remove') {
                if (!confirm(`移除「${name}」的權限？\n只有他不能再用，其他人不用改。他的手機下次連線時，App 裡的公司資料會自動清除。`)) return
                await postSync({ action: 'remove', id })
                toast(`已移除「${name}」`)
              }
              load()
            } catch (e) {
              toast(e.message)
            }
          }),
        )
      }
      load()
      el.querySelector('#sh-invite').onclick = async () => {
        const name = el.querySelector('#sh-name').value.trim()
        if (!name) return toast('請先填名字')
        const role = el.querySelector('#sh-role').value
        const key = newSyncKey()
        try {
          await postSync({ action: 'invite', name, role, newKey: key })
        } catch (e) {
          return toast(e.message)
        }
        close()
        linkSheet(name, role, `${ls.get(LS.sheet)}#k=${key}`, appUrl)
      }
      el.querySelector('#sh-move')?.addEventListener('click', () => {
        close()
        moveSheet()
      })
    },
  )
}
/** 選盤點人：共用名單裡的人（名字由擁有者／管理員取） */
function counterSheet(currentId, title, sub, onPick) {
  const people = roster()
  if (!people.length) return toast('同步一次後才有名單；先按「立即同步」')
  sheet(
    `<h2 class="sheet-title">${esc(title)}</h2><p class="sheet-sub">${esc(sub)}</p>
     <div class="group">${people
       .map((p) => `<button class="row" data-pick="${esc(p.id)}" ${p.id === currentId ? 'aria-current="true"' : ''}><span class="avatar" aria-hidden="true">${esc(p.name.slice(0, 1))}</span><span class="grow"><span class="title">${esc(p.name)}</span></span>${p.id === currentId ? '<span class="muted">✓</span>' : ''}</button>`)
       .join('')}</div>
     <p class="footnote">名單裡沒有的人：到「共用設定」邀請，名字由擁有者或管理員取。</p>`,
    (el, close) =>
      el.querySelectorAll('[data-pick]').forEach((b) =>
        b.addEventListener('click', async () => {
          close()
          await onPick(people.find((p) => p.id === b.dataset.pick))
        }),
      ),
  )
}
/** 邀請完：顯示只給這個人的連結碼（關掉就看不到了，雲端只存雜湊值） */
function linkSheet(name, role, code, appUrl) {
  const text = `拍照盤點的連結碼（只給你用，不要轉傳）：\n${code}\n\n打開 ${appUrl} → 右上「設定」→「我收到連結碼了」→ 貼上 → 加入`
  sheet(
    `<h2 class="sheet-title">給「${esc(name)}」的連結碼</h2>
     <p class="sheet-sub">${ROLE_LABEL[role]}：${ROLE_DESC[role]}</p>
     <textarea class="field code" readonly rows="3">${esc(code)}</textarea>
     <div class="row-actions" style="margin-top:10px"><button class="btn" id="lk-copy" style="flex:1">複製</button>${navigator.share ? '<button class="btn secondary" id="lk-share" style="flex:1">傳給他…</button>' : ''}</div>
     <ol class="steps-list">
       <li>私訊給「${esc(name)}」（不要貼在群組）。</li>
       <li>他打開拍照盤點 →「設定」→「我收到連結碼了」→ 貼上 →「加入」。</li>
       <li>這組只給他一個人用。他離職時，在「共用設定」移除權限就好，其他人不用改。</li>
     </ol>
     <p class="footnote">關掉之後就看不到這組連結碼了（雲端只存雜湊值，比較安全）；忘了可以「重新產生連結碼」。</p>`,
    (el) => {
      el.querySelector('#lk-copy').onclick = async () => {
        try {
          await navigator.clipboard.writeText(text)
          toast('已複製（含使用說明）：私訊給他')
        } catch {
          el.querySelector('textarea').select()
          toast('請長按上面的連結碼複製')
        }
      }
      el.querySelector('#lk-share')?.addEventListener('click', () => navigator.share({ text }).catch(() => {}))
    },
  )
}
/** 改名字（擁有者自己 id＝owner） */
function nameSheet(id, title, sub) {
  sheet(
    `<h2 class="sheet-title">${esc(title)}</h2>${sub ? `<p class="sheet-sub">${esc(sub)}</p>` : ''}
     <input class="field" id="nm-val" value="${esc(id === (ls.get(LS.memberId) || 'owner') ? ls.get(LS.memberName) : '')}" maxlength="30" autocomplete="off">
     <button class="btn block" id="nm-save" style="margin-top:12px">儲存</button>`,
    (el, close) => {
      const input = el.querySelector('#nm-val')
      input.focus()
      el.querySelector('#nm-save').onclick = async () => {
        const name = input.value.trim()
        if (!name) return toast('名字不能空白')
        try {
          await postSync({ action: 'rename', id, name })
        } catch (e) {
          return toast(e.message)
        }
        if (id === (ls.get(LS.memberId) || 'owner')) ls.set(LS.memberName, name)
        close()
        toast('已儲存')
        render()
      }
    },
  )
}
/** 搬到另一個 Google 帳號（只有擁有者）：資料、共用的人、權限一起搬；舊位置會叫大家自動跟過去 */
function moveSheet() {
  sheet(
    `<h2 class="sheet-title">搬到另一個 Google 帳號</h2>
     <p class="sheet-sub">例如從個人帳號搬到公司帳號。資料、共用的人和權限都會一起搬；大家的 App 會自動跟過去，不用重貼連結碼。</p>
     <ol class="steps-list">
       <li>用新的帳號開一個 Google 試算表 →「擴充功能」→「Apps Script」→ 貼上程式碼（設定 →「複製試算表程式碼」）→ 部署成網頁應用程式（執行身分：我；存取：所有人）。</li>
       <li>把新的「網頁應用程式網址」貼在下面，按「開始搬家」。這支手機要開著、有網路，直到顯示完成。</li>
     </ol>
     <input class="field" id="mv-url" placeholder="https://script.google.com/macros/s/…/exec" autocomplete="off" spellcheck="false">
     <button class="btn block" id="mv-go" style="margin-top:12px">開始搬家</button>
     <p class="footnote" id="mv-status"></p>`,
    (el) => {
      const status = (m) => (el.querySelector('#mv-status').textContent = m)
      el.querySelector('#mv-go').onclick = async (ev) => {
        const newUrl = el.querySelector('#mv-url').value.trim()
        const oldUrl = ls.get(LS.sheet)
        if (!/^https:\/\/script\.google\.com\/macros\/s\/[^#\s]+\/exec$/.test(newUrl)) return toast('網址不對：要是 https://script.google.com/macros/s/…/exec')
        if (newUrl === oldUrl) return toast('這是現在用的網址；請貼新帳號的網址')
        const btn = ev.currentTarget
        btn.disabled = true
        state.moving = true // 搬家中：背景自動同步先停（網址會暫時切來切去）
        clearTimeout(syncTimer)
        try {
          status('1／4 先把最新的資料下載到這台…')
          await syncNow()
          status('2／4 在新位置開啟同步…')
          ls.set(LS.sheet, newUrl)
          await postSync({ action: 'hello' })
          ls.set(LS.sheet, oldUrl)
          const ex = await postSync({ action: 'exportMembers' })
          ls.set(LS.sheet, newUrl)
          await postSync({ action: 'importMembers', members: ex.members, revokedHashes: ex.revokedHashes, ownerName: ex.ownerName })
          status('3／4 把全部資料傳到新位置（照片多會比較久）…')
          for (const store of Object.values(SYNC_STORES)) for (const v of await store.all()) await store.putRaw({ ...v, _syncT: undefined })
          ls.set(LS.settingsSyncT, '')
          ls.set(LS.pulled, '0')
          itemsCache = null
          await syncNow(status)
          status('4／4 通知大家改用新位置…')
          ls.set(LS.sheet, oldUrl)
          await postSync({ action: 'moveTo', url: newUrl })
          ls.set(LS.sheet, newUrl)
          status('搬家完成：大家的 App 下次同步會自動跟過來。舊的試算表和資料夾確認沒問題後可以刪掉。')
          toast('搬家完成')
        } catch (e) {
          ls.set(LS.sheet, oldUrl)
          status(`沒有完成：${e?.message || String(e)}。資料還在原本的位置，可以再試一次。`)
          btn.disabled = false
        } finally {
          state.moving = false
        }
      }
    },
  )
}
const syncLabel = () => {
  if (state.syncState === 'syncing') return '同步中…'
  if (state.syncState === 'error') return '同步沒成功'
  const t = Number(ls.get(LS.lastSync, '0'))
  return t ? `已同步 ${fmtTime(t)}` : '還沒同步'
}

// ───────────────────────── 總表：很多次盤點合在一起 ─────────────────────────
const pad2 = (n) => String(n).padStart(2, '0')
const ymd = (t) => {
  const d = new Date(t)
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`
}
const hm = (t) => {
  const d = new Date(t)
  return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`
}
const RANGES = [
  { id: 'today', label: '今天' },
  { id: 'week', label: '最近 7 天' },
  { id: 'all', label: '全部' },
]
function inRange(s, range) {
  if (range === 'all') return true
  const start = new Date()
  start.setHours(0, 0, 0, 0)
  if (range === 'week') start.setDate(start.getDate() - 6)
  return s.createdAt >= start.getTime()
}

/**
 * 把好幾次盤點合起來：同一種商品（品名＋品牌＋型號＋尺寸都一樣）的數量加總，記下在哪些位置各幾件。
 * detail＝每次盤點、每一種一列（給 Excel 明細、Google 試算表用）。
 */
function reportOf(sessions) {
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  const items = new Map()
  const detail = []
  for (const s of [...sessions].sort((a, b) => a.createdAt - b.createdAt)) {
    const place = s.place || `未填位置（${fmtTime(s.createdAt)}）`
    for (const g of groupsOf(s)) {
      const count = Number(g.count) || 0
      if (!count) continue
      const k = FIELDS.map((f) => norm(g[f])).join('|')
      if (!items.has(k)) items.set(k, { label: g.label, brand: g.brand, model: g.model, spec: g.spec, total: 0, places: new Map() })
      const it = items.get(k)
      it.total += count
      it.places.set(place, (it.places.get(place) || 0) + count)
      detail.push({ date: ymd(s.createdAt), time: hm(s.createdAt), place: s.place || '', label: g.label, brand: g.brand, model: g.model, spec: g.spec, count, boxes: g.manual ? '' : g.boxes, source: g.manual ? '手動' : g.edited ? 'AI（有修正）' : 'AI', id: s.id, by: byName(s) })
    }
  }
  const list = [...items.values()].sort((a, b) => cmp(a.label, b.label) || cmp(a.spec, b.spec) || cmp(a.brand, b.brand))
  const places = new Set(sessions.map((s) => s.place || s.id))
  return { list, detail, total: list.reduce((n, it) => n + it.total, 0), places: places.size, sessions: sessions.length }
}
const placesText = (it) => [...it.places].map(([p, n]) => `${p} ${n}`).join('、')

/** Excel：總表＋明細＋盤點清單（＋品項庫、叫貨清單） */
function reportXlsx(sessions, items = [], whoOf = () => '') {
  const r = reportOf(sessions)
  return makeXlsx([
    { name: '總表', rows: [['品名', '尺寸／規格', '品牌', '型號', '總數量', '在哪裡（位置 數量）'], ...r.list.map((it) => [it.label, it.spec, it.brand, it.model, it.total, placesText(it)])] },
    { name: '明細', rows: [['盤點日期', '時間', '盤點人', '位置', '品名', '品牌', '型號', '尺寸／規格', '數量', '照片框數', '來源'], ...r.detail.map((d) => [d.date, d.time, d.by, d.place, d.label, d.brand, d.model, d.spec, d.count, d.boxes, d.source])] },
    {
      name: '盤點清單',
      rows: [['盤點時間', '盤點人', '位置', '照片張數', '種類', '件數', '辨識模型'], ...[...sessions].sort((a, b) => a.createdAt - b.createdAt).map((s) => [`${ymd(s.createdAt)} ${hm(s.createdAt)}`, byName(s), s.place || '', s.photos.length, groupsOf(s).length, totalQty(s), s.model || ''])],
    },
    ...(items.length ? itemSheets(items, whoOf) : []),
  ])
}
/** 複製成表格（Tab 分隔）：在 Google 試算表或 Excel 點一格、貼上，就會自動分好欄 */
const reportTsv = (sessions) => {
  const r = reportOf(sessions)
  const clean = (v) => String(v ?? '').replace(/[\t\n]/g, ' ')
  return [['品名', '尺寸／規格', '品牌', '型號', '總數量', '在哪裡（位置 數量）'], ...r.list.map((it) => [it.label, it.spec, it.brand, it.model, it.total, placesText(it)])].map((row) => row.map(clean).join('\t')).join('\n')
}

/**
 * 同步到 Google 試算表（使用者自己的 Apps Script 網頁應用程式）。
 * 同一次盤點重送不會重複（試算表那邊會先刪掉同一個盤點 ID 的舊列）。
 * 有些情況瀏覽器讀不到回覆（跨網域）：改成「只送出、不看回覆」，回傳 null＝已送出但無法確認。
 */
async function syncToSheet(sessions) {
  const url = ls.get(LS.sheet)
  if (!url) throw new Error('還沒設定 Google 試算表')
  const rows = reportOf(sessions).detail.map((d) => [d.date, d.time, d.place, d.label, d.brand, d.model, d.spec, d.count, d.boxes, d.source, d.id, d.by])
  // 報表資料整份一起送（試算表那邊重寫「總覽、庫存、儲位庫存、盤差報告」；舊版試算表程式碼會忽略）
  const report = await sheetReport()
  const body = JSON.stringify({ rows, report, key: ls.get(LS.syncKey) || undefined, dev: deviceId() })
  let confirmed = null
  let data = null
  try {
    data = await (await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body })).json()
  } catch {
    data = null
  }
  // 試算表有回話但說不行（例如同步密碼不對）→ 照實說；讀不到回覆 → 改成「只送出、不看回覆」
  if (data && !data.ok) throw new Error(data.error || '試算表回傳錯誤')
  if (data) confirmed = data.rows
  else await fetch(url, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body })
  const now = Date.now()
  for (const s of sessions) {
    s.syncedAt = now
    await db.putRaw(s) // 只記本機狀態：不要因此把照片再傳一次到雲端
  }
  return confirmed
}

function downloadBlob(blob, name) {
  const a = document.createElement('a')
  a.href = URL.createObjectURL(blob)
  a.download = name
  a.click()
  setTimeout(() => URL.revokeObjectURL(a.href), 4000)
}

async function viewReport() {
  const all = await db.all()
  const range = state.range || 'today'
  const sessions = all.filter((s) => inRange(s, range))
  const r = reportOf(sessions)
  const hasSheet = !!ls.get(LS.sheet)
  return `
  <main class="app">
    <div class="nav">${backBtn('home', '盤點')}</div>
    <h1 class="large-title">總表</h1>
    <p class="subtitle">很多次盤點合在一起：同一種商品自動加總，看得到在哪裡各幾件。</p>
    <div class="seg" role="tablist" aria-label="時間範圍">${RANGES.map((x) => `<button role="tab" aria-selected="${x.id === range}" data-range="${x.id}">${x.label}</button>`).join('')}</div>
    ${
      sessions.length
        ? `<div class="stats">
            <div class="stat"><span class="stat-num">${r.total}</span><span class="stat-label">件</span></div>
            <div class="stat"><span class="stat-num">${r.list.length}</span><span class="stat-label">種商品</span></div>
            <div class="stat"><span class="stat-num">${r.places}</span><span class="stat-label">個位置</span></div>
            <div class="stat"><span class="stat-num">${r.sessions}</span><span class="stat-label">次盤點</span></div>
          </div>
          <div class="row-actions" style="margin:14px 0 6px">
            <button class="btn" style="flex:1" data-action="report-xlsx">下載 Excel</button>
            <button class="btn secondary" style="flex:1" data-action="report-copy">複製表格</button>
          </div>
          ${
            hasSheet
              ? `<button class="btn secondary block" data-action="report-sync">同步到 Google 試算表（${r.sessions} 次盤點）</button>`
              : `<button class="btn plain block" data-go="settings">連結 Google 試算表…</button><p class="footnote" style="margin-top:6px">連結後，每次盤點完會自動寫進你的試算表。</p>`
          }
          <p class="footnote">Excel 有：總表、明細（每次盤點每一種一列）、盤點清單，還有品項庫（實盤、帳面、差異）和叫貨清單。「複製表格」後，在 Google 試算表點一格貼上，就會自動分好欄。</p>
          <p class="section-title">品項（${r.list.length}）</p>
          <div class="group">${r.list
            .map(
              (it) => `<div class="row report-row"><span class="grow"><span class="title">${esc(it.label)}</span><br><span class="meta">${esc([it.spec, it.brand, it.model].filter(Boolean).join('・') || '沒有寫尺寸')}</span><br><span class="meta">${esc(placesText(it))}</span></span><span class="report-count">${it.total}</span></div>`,
            )
            .join('')}</div>`
        : `<div class="empty"><p>${range === 'today' ? '今天還沒有盤點。' : '這段時間沒有盤點紀錄。'}<br>換一個時間範圍看看。</p></div>`
    }
  </main>`
}

// ───────────────────────── 匯出 ─────────────────────────
function csvOf(session) {
  const cell = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`
  const rows = [['盤點時間', '位置', '品名', '品牌', '型號', '尺寸／規格', '數量', '照片框數', 'AI 信心', '來源']]
  for (const g of groupsOf(session)) rows.push([fmtTime(session.createdAt), session.place, g.label, g.brand, g.model, g.spec, g.count, g.manual ? '' : g.boxes, g.manual ? '' : g.conf.toFixed(2), g.manual ? '手動' : g.edited ? 'AI（有修正）' : 'AI'])
  return '﻿' + rows.map((r) => r.map(cell).join(',')).join('\r\n') // 加 BOM，Excel 開才不會亂碼
}
const textOf = (session) => [`盤點 ${fmtTime(session.createdAt)}${session.place ? `・${session.place}` : ''}`, ...groupsOf(session).map((g) => `・${g.label}${detailOf(g) ? `（${detailOf(g)}）` : ''}：${g.count}`), `共 ${totalQty(session)} 件`].join('\n')

function exportSheet() {
  const s = state.session
  const name = `盤點_${(s.place || summaryOf(s)).replace(/[\\/:*?"<>|\s]+/g, '_')}_${new Date(s.createdAt).toISOString().slice(0, 10)}.csv`
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
  await itemsAll() // 品項庫：給 AI 照抄名稱、填料號
  const status = (msg) => {
    const el = document.getElementById('ai-status')
    if (el) el.textContent = msg + (refs.length ? `（附 ${refs.length} 張樣品照）` : '')
  }
  // 多張照片一次送兩張（總時間差不多減半；再多會撞到免費額度每分鐘上限）
  const queue = [...indices]
  const worker = async () => {
    while (queue.length && !state.cancel) {
      const photo = s.photos[queue.shift()]
      // 辨識過程（用哪個模型、花幾秒、哪一步失敗）記在照片上，結果頁可以展開看、截圖回報
      photo.trace = []
      const t0 = Date.now()
      const log = (msg) => {
        photo.trace.push(`${Math.round((Date.now() - t0) / 100) / 10} 秒｜${msg}`)
        status(msg)
      }
      try {
        const r = await analyze(photo, log, refs)
        Object.assign(photo, { objects: r.objects, note: r.note, status: 'done', error: '', errorDetail: '' })
        s.model = r.model
        // 拍到儲位標籤：沒填位置就自動填（有建立過的儲位用那個寫法）
        if (r.location && !s.place) {
          s.place = findLocation(r.location)?.code || r.location.toUpperCase()
          s.placeAuto = true
          log(`看到儲位標籤 ${s.place}，自動填位置`)
        }
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
  if (s.placeAuto) {
    delete s.placeAuto
    toast(`看到儲位標籤，位置自動填「${placeLabel(s.place)}」`)
  }
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
  let failed = 0
  for (const { photo, objs } of jobs) {
    const free = objs.filter((o) => !o.edited && photo.objects.includes(o))
    const t0 = Date.now()
    const log = (msg) => (photo.trace ??= []).push(`比對 ${Math.round((Date.now() - t0) / 100) / 10} 秒｜${msg}`)
    try {
      const rr = await refineObjects(photo, free, free.map((_, k) => k), refs, log)
      groups += rr.groups || 0
      if (rr.model && !/比對/.test(s.model || '')) s.model = rr.model === s.model ? `${s.model}（含相似品比對）` : `${s.model} ＋ ${rr.model} 比對`
    } catch (e) {
      failed += 1
      log(`沒成功：${e.detail || e.message}`)
      photo.note = [photo.note, 'AI 沒辦法分出相似品的尺寸：請點照片上的框，直接選它是哪一種。'].filter(Boolean).join(' ')
    }
    state.refining.done += 1
    await db.put(s)
    if (state.session === s && state.view === 'review') render()
  }
  state.refining = null
  if (state.session === s && state.view === 'review') {
    render()
    toast(failed ? '相似品沒有比對成功：點照片上的框，直接選它是哪一種' : groups > 1 ? '相似品比對完成，已經分開不同尺寸' : '相似品比對完成：AI 認為都是同一種')
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
  // 補框模式：點照片哪裡，就在那裡加一個框（大小跟這張照片的其他框差不多），再選它是哪一種
  const wrap = e.target.closest('[data-photo]')
  if (state.addMode && wrap && !e.target.closest('[data-action]')) return addBoxAt(wrap, e)
  const t = e.target.closest('[data-go],[data-action],[data-open],[data-remove-photo],[data-photo-index],[data-focus],[data-step],[data-edit],[data-retry],[data-range],[data-item-open],[data-item-filter],[data-place],[data-loc-edit],[data-stock-del],[data-safety],[data-example]')
  if (!t) {
    // 點空白處取消標示（不捲回頂端）
    if (state.focus && !e.target.closest('.photo-wrap,.item')) {
      state.focus = null
      render()
    }
    return
  }
  const d = t.dataset
  // 檢視者：擋掉所有會改資料的動作（按鈕本來就藏起來了，這裡再保險一次）
  if (!canEdit() && (EDIT_ACTIONS.has(d.action) || d.step || d.edit || d.retry !== undefined || d.locEdit !== undefined || d.stockDel || d.safety || d.removePhoto !== undefined || d.place !== undefined)) return toast('你是檢視者：只能看，不能修改')
  if (!canEdit() && d.focus !== undefined && d.obj) {
    state.focus = state.focus === d.focus ? null : d.focus
    return render()
  }
  if (d.go) return go(d.go)
  if (d.range) {
    state.range = d.range
    return render()
  }
  if (d.open) {
    state.session = await db.get(d.open)
    delete state.session.edits // 第一版的舊欄位，不再使用
    await itemsAll()
    return go('review', { photoIndex: 0 })
  }
  if (d.itemOpen) {
    state.itemId = d.itemOpen
    return go('item')
  }
  if (d.itemFilter) {
    state.itemFilter = d.itemFilter
    return render()
  }
  if (d.place !== undefined) {
    state.session.place = canon(state.session.place) === canon(d.place) ? '' : d.place
    return render()
  }
  if (d.locEdit !== undefined) return locationSheet(Number(d.locEdit))
  if (d.example) {
    state.lookup = { q: d.example, read: null, busy: false }
    return render()
  }
  if (d.stockDel) {
    const it = await currentItem()
    const st = it?.stock[d.stockDel]
    if (!st || st.removed || !confirm(`拿掉「${stockPlace(st)}」的 ${st.count} 個？（盤錯位置、重複算時用）`)) return
    it.stock[d.stockDel] = removedEntry(st)
    await putItem(it)
    return render()
  }
  if (d.safety) {
    const it = await currentItem()
    if (!it) return
    it.safety = Math.max(0, (it.safety ?? 0) + Number(d.safety))
    await putItem(it)
    return render()
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
    // 點照片上的框＝「這一個是哪一種？」直接改（補框模式時，點到框也當成補在那裡）
    if (d.obj && !state.addMode) {
      const [pi, oi] = d.obj.split(':').map(Number)
      const entries = photoEntries(pi)
      const doubts = doubtsOf(state.session)
      const target = state.session.photos[pi].objects[oi]
      // 帶上「為什麼要看」的理由
      entries.forEach((en) => (en.reason = doubts.find((x) => x.o === en.o)?.reason || ''))
      return quickSheet(entries, entries.findIndex((en) => en.o === target))
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
      // 盤點人＝這台選好的人（不用每次選；擁有者／管理員可以在設定改）
      state.session = { id: uid(), createdAt: Date.now(), place: '', photos: [], counts: {}, manual: [], by: currentCounter()?.name || '', byId: currentCounter()?.id || '' }
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
      if (!confirm('確定刪除這次盤點？照片和結果都會刪掉，記在品項庫的數量也會拿掉。')) return
      await db.del(state.session.id)
      tombstone(`session:${state.session.id}`)
      await unlinkSession(state.session.id)
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
      touchSettings()
      return toast('清單已儲存')
    case 'reset-catalog':
      ls.set(LS.catalog, '')
      touchSettings()
      toast('已恢復預設清單')
      return render()
    case 'force-update':
      return forceUpdate()
    // ── 多台裝置同步 ──
    case 'sync-start': {
      ls.set(LS.syncKey, newSyncKey())
      try {
        await postSync({ action: 'hello' })
      } catch (err) {
        ls.set(LS.syncKey, '')
        return toast(err.message)
      }
      toast('已開啟同步，第一次上傳中…（照片多會比較久）')
      render()
      syncNow((m) => toast(m))
        .then((r) => toast(`同步完成：上傳 ${r.pushed} 筆。接著到「共用設定」邀請同事`))
        .catch((err) => toast(`同步沒成功：${err.message}`))
        .finally(() => state.view === 'settings' && render())
      // 先問名字：大家會看到「誰盤的」
      return nameSheet('owner', '你的名字？', '大家會在盤點紀錄看到是誰盤的（例如：小宇、店長）。')
    }
    case 'sync-link': {
      const code = parseLinkCode(document.getElementById('link-code').value)
      if (!code) return toast('連結碼不對：要整段貼上（https://script.google.com/…/exec#k=…）')
      const before = { url: ls.get(LS.sheet), key: ls.get(LS.syncKey) }
      ls.set(LS.sheet, code.url)
      ls.set(LS.syncKey, code.key)
      try {
        await postSync({ action: 'hello' })
      } catch (err) {
        ls.set(LS.sheet, before.url)
        ls.set(LS.syncKey, before.key)
        return toast(err.message)
      }
      ls.set(LS.pulled, '0') // 新連結：從頭下載一次
      toast('連結成功，下載資料中…')
      render()
      return syncNow((m) => toast(m))
        .then((r) => toast(`同步完成：下載 ${r.pulled} 筆、上傳 ${r.pushed} 筆`))
        .catch((err) => toast(`同步沒成功：${err.message}`))
        .finally(() => render())
    }
    case 'share-open':
      return shareSheet()
    case 'counter-device':
    case 'counter-session': {
      if (!canManage()) return toast('只有擁有者和管理員可以改盤點人')
      const forSession = d.action === 'counter-session'
      const s = state.session
      const cur = forSession ? s?.byId : currentCounter()?.id
      return counterSheet(cur, forSession ? '這次是誰盤的？' : '這台的盤點人', forSession ? '盤點時間不會變。' : '之後在這台按「新盤點」，都會記成這個人（例如店裡共用的平板）。', async (p) => {
        if (forSession) {
          Object.assign(s, { by: p.name, byId: p.id })
          await save()
          toast(`盤點人改成「${p.name}」`)
        } else {
          ls.set(LS.counterId, p.id === ls.get(LS.memberId) ? '' : p.id)
          toast(`這台的盤點人：${p.name}`)
        }
        render()
      })
    }
    case 'sync-leave':
      if (!confirm('這台退出同步，並清除這台的盤點紀錄、品項庫、樣品照、儲位？\n雲端和其他人的資料不會刪。API Key 也會一起清掉。')) return
      return wipeLocal('這台已退出同步，資料已清除')
    case 'sync-now':
      if (!syncReady()) return go('settings')
      toast('同步中…')
      return syncNow((m) => toast(m))
        .then((r) => toast(r.pushed || r.pulled ? `同步完成：上傳 ${r.pushed}、下載 ${r.pulled}` : '已經是最新的'))
        .catch((err) => toast(`同步沒成功：${err.message}`))
        .finally(() => ['home', 'settings'].includes(state.view) && render())
    case 'finish': {
      // 完成：記進品項庫；有設定 Google 試算表又開著自動同步，就順便寫進去（在背景送，不用等）
      const s = state.session
      if (s) {
        t.disabled = true
        const r = await linkSession(s).catch(() => null)
        toast(r ? (r.created ? `已記進品項庫：${r.linked} 種（新的 ${r.created} 種，到「品項」確認名稱）` : `已記進品項庫：${r.linked} 種`) : '記進品項庫失敗：請打開這次盤點，再按一次完成')
      }
      go('home')
      // 盤完馬上傳給大家（不用等）
      if (syncReady()) syncNow().catch((err) => toast(`同步沒成功：${err.message}；有網路時會再自動試`))
      if (s && ls.get(LS.sheet) && ls.get(LS.autoSync, '1') === '1') {
        syncToSheet([s])
          .then((n) => {
            toast(n === null ? '已送到 Google 試算表' : `已同步到 Google 試算表（${n} 列）`)
            if (state.view === 'home') render()
          })
          .catch((err) => toast(`同步沒成功：${err.message}；可以到「總表」再按一次同步`))
      }
      return
    }
    case 'report-xlsx': {
      const sessions = (await db.all()).filter((s) => inRange(s, state.range || 'today'))
      downloadBlob(reportXlsx(sessions, await itemsAll(true), await whoOfStock()), `盤點總表_${ymd(Date.now())}.xlsx`)
      return toast('已下載 Excel')
    }
    case 'report-copy': {
      const sessions = (await db.all()).filter((s) => inRange(s, state.range || 'today'))
      try {
        await navigator.clipboard.writeText(reportTsv(sessions))
        return toast('已複製：到 Google 試算表點一格，貼上就好')
      } catch {
        return toast('這個瀏覽器不讓複製，請改用「下載 Excel」')
      }
    }
    case 'report-sync': {
      const sessions = (await db.all()).filter((s) => inRange(s, state.range || 'today'))
      t.disabled = true
      t.textContent = '同步中…'
      try {
        const n = await syncToSheet(sessions)
        toast(n === null ? `已送出 ${sessions.length} 次盤點，請打開試算表確認` : `已同步 ${sessions.length} 次盤點（${n} 列）`)
      } catch (err) {
        toast(`同步沒成功：${err.message}`)
      }
      return render()
    }
    case 'sheet-save': {
      const url = document.getElementById('sheet-url').value.trim()
      if (url && !/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec/.test(url)) return toast('網址不對：要是 https://script.google.com/macros/s/…/exec')
      ls.set(LS.sheet, url)
      if (!url) return toast('已取消連結 Google 試算表')
      try {
        const data = await (await fetch(url)).json()
        toast(data.ok ? `連上了：${data.sheet || 'Google 試算表'}` : '連得到，但回覆不對；請確認貼的是這個 App 的程式碼')
      } catch {
        toast('已儲存；測試時瀏覽器讀不到回覆，同步一次後打開試算表確認')
      }
      return
    }
    case 'sheet-copy':
      try {
        const code = await (await fetch('google-sheets.gs', { cache: 'no-cache' })).text()
        await navigator.clipboard.writeText(code)
        return toast('已複製程式碼：到 Apps Script 貼上')
      } catch {
        return toast('複製失敗，請改用電腦打開這個 App 再按一次')
      }
    case 'review-doubts':
      return quickSheet(doubtsOf(state.session), 0, { doubt: true })
    case 'add-box':
      state.addMode = true
      state.focus = null
      render()
      return toast('點照片上漏掉的那一個')
    case 'add-cancel':
      state.addMode = false
      return render()
    case 'zoom':
      state.viewer = true
      state.zoom = 2
      return render()
    case 'zoom-in':
      state.zoom = Math.min(4, state.zoom + 1)
      return render()
    case 'zoom-out':
      state.zoom = Math.max(1, state.zoom - 1)
      return render()
    case 'zoom-close':
      state.viewer = false
      return render()
    case 'del-sample':
      if (!confirm('刪掉這張樣品照？')) return
      await idb.samples.del(d.id)
      tombstone(`sample:${d.id}`)
      toast('已刪掉樣品照')
      return render()
    case 'clear-all':
      if (!confirm(`確定刪除全部盤點紀錄（含照片）？刪了救不回來。\n品項庫、各位置的數量、儲位會保留。${syncReady() ? '\n有開多台同步：其他裝置的盤點紀錄也會一起刪掉。' : ''}`)) return
      for (const s of await db.all()) tombstone(`session:${s.id}`)
      await db.clear()
      return toast('已全部刪除；品項庫保留')
    // ── 品項庫 ──
    case 'item-add':
      return itemEditSheet(null)
    case 'items-more':
      return itemsMoreSheet()
    case 'import':
      return importSheet()
    case 'item-edit':
      return itemEditSheet(await currentItem())
    case 'item-confirm': {
      const it = await currentItem()
      if (!it) return
      it.status = 'ok'
      await putItem(it)
      toast('已確認：以後辨識會照這個名稱寫')
      return render()
    }
    case 'item-merge': {
      const it = await currentItem()
      if (!it) return
      return pickItem('合併到哪一項？', `「${itemTitle(it)}」的數量、以前的寫法會併過去，這一項會刪掉。`, [it], async (to) => {
        if (!to || !confirm(`把「${itemTitle(it)}」合併到「${itemTitle(to)}」（${to.no}）？`)) return
        await mergeItems(it, to)
        state.itemId = to.id
        go('item')
        toast(`已合併到 ${to.no}`)
      })
    }
    case 'item-delete': {
      const it = await currentItem()
      if (!it || !confirm(`刪除「${itemTitle(it)}」（${it.no}）？各位置的數量、進出紀錄都會刪掉。`)) return
      await delItem(it)
      toast('已刪除')
      return go('items')
    }
    case 'move-in':
    case 'move-out': {
      const it = await currentItem()
      if (!it) return
      const inbound = d.action === 'move-in'
      return numberSheet({ title: inbound ? '進貨幾個？' : '賣出幾個？', sub: `${esc(itemTitle(it))}：帳面數${inbound ? '加' : '減'}這麼多。${it.book == null ? `還沒有帳面數，會從實盤 ${onHand(it)} 開始算。` : `目前帳面 ${it.book}。`}`, value: '1', action: inbound ? '記進貨' : '記賣出' }, async (n) => {
        const now = Date.now()
        // 還沒有帳面數：先記一筆「從實盤開始」，帳面數才算得回來（多人同步時用進出紀錄重算）
        if (it.book == null) it.moves.push({ at: now - 1, kind: 'set', qty: onHand(it) })
        it.moves.push({ at: now, kind: inbound ? 'in' : 'out', qty: n })
        it.book = bookFromMoves(it.moves)
        await putItem(it)
        toast(`帳面數變成 ${it.book}`)
      })
    }
    case 'book-set': {
      const it = await currentItem()
      if (!it) return
      return numberSheet({ title: '帳面數', sub: '應該要有幾個。第一次盤點完，直接用實盤數當起點最快；之後進貨、賣出再加減。', value: it.book ?? onHand(it), quick: liveStock(it).length ? [{ n: onHand(it), label: `用實盤數 ${onHand(it)}` }] : [] }, async (n) => {
        it.book = n
        it.moves.push({ at: Date.now(), kind: 'set', qty: n })
        await putItem(it)
      })
    }
    case 'equiv-add': {
      const it = await currentItem()
      if (!it) return
      return pickItem('哪一項可以互換？', '設定後兩邊都會顯示，客人問的時候一查就知道。', [it], async (x) => {
        if (!x) return
        it.equiv = [...new Set([...(it.equiv || []), x.id])]
        x.equiv = [...new Set([...(x.equiv || []), it.id])]
        await putItem(it)
        await putItem(x)
        render()
      })
    }
    case 'read-add': {
      const r = state.lookup.read
      if (!r) return
      if (!r.label) return toast('AI 看不出品名：請按清除後打型號，或到品項庫按 ＋ 自己填')
      const items = await itemsAll()
      const dup = findItem(items, r)
      if (dup) {
        state.itemId = dup.id
        return go('item')
      }
      const it = newItem(items, r, { status: 'ok', photo: r.thumb })
      await putItem(it)
      // 也存成樣品照：以後盤點時 AI 會拿來比對
      await idb.samples.put({ id: uid(), label: r.label, brand: r.brand, model: r.model, spec: r.spec, blob: r.thumb, createdAt: Date.now() })
      toast(`已加入品項庫 ${it.no}，也存成樣品照`)
      state.itemId = it.id
      return go('item')
    }
    case 'read-clear':
      if (state.lookup.read?.url) URL.revokeObjectURL(state.lookup.read.url)
      state.lookup = { q: '', read: null, busy: false }
      return render()
    case 'safety-pick':
      // 「叫貨」是空的：直接選一個品項 → 填「剩幾個就要叫貨」
      return pickItem('設定哪一種？', '選一個品項，再填「剩幾個就要叫貨」。之後可以一個一個設。', [], async (it) => {
        if (!it) return
        numberSheet({ title: '剩幾個就要叫貨？', sub: `${esc(itemTitle(it))}：現在大概剩 ${expected(it)} 個。填 3 的意思是：剩 3 個以下就提醒叫貨。`, value: it.safety ?? 3, action: '儲存' }, async (n) => {
          it.safety = n
          await putItem(it)
          toast(needsOrder(it) ? `已設定：${itemTitle(it)} 現在就該叫貨了` : `已設定：剩 ${n} 個以下會提醒叫貨`)
        })
      })
    case 'order-copy': {
      const list = (await itemsAll()).filter(needsOrder)
      try {
        await navigator.clipboard.writeText(orderText(list))
        return toast(`已複製 ${list.length} 項叫貨清單，到 LINE 貼上`)
      } catch {
        return toast('這個瀏覽器不讓複製，請改用 ⋯ → 下載品項庫 Excel')
      }
    }
    case 'loc-add':
      return locationSheet(-1)
    case 'loc-print':
      return printLabels()
  }
})

function bindInputs() {
  document.getElementById('cam')?.addEventListener('change', (e) => addFiles([...e.target.files]))
  document.getElementById('pick')?.addEventListener('change', (e) => addFiles([...e.target.files]))
  document.getElementById('sample-cam')?.addEventListener('change', (e) => e.target.files[0] && newSampleSheet(e.target.files[0]))
  document.getElementById('auto-sync')?.addEventListener('change', (e) => ls.set(LS.autoSync, e.target.checked ? '1' : '0'))
  document.getElementById('place')?.addEventListener('input', (e) => (state.session.place = e.target.value))
  document.getElementById('item-search')?.addEventListener('input', (e) => filterRows($app, e.target.value))
  document.getElementById('lookup-q')?.addEventListener('input', (e) => {
    // 只更新結果區，輸入框不重畫（打字、選字不會被打斷）
    if (state.lookup.read?.url) URL.revokeObjectURL(state.lookup.read.url)
    state.lookup = { q: e.target.value.trim(), read: null, busy: false }
    document.getElementById('lookup-results').innerHTML = lookupResults()
  })
  document.getElementById('label-cam')?.addEventListener('change', async (e) => {
    const file = e.target.files[0]
    if (!file) return
    state.lookup = { q: '', read: null, busy: true }
    render()
    try {
      const read = await readLabel(file)
      state.lookup = { q: read.model || read.label, read, busy: false }
    } catch (err) {
      state.lookup = { q: '', read: null, busy: false }
      toast(err.message || 'AI 沒讀到，請再拍一次（靠近一點、不要反光）')
    }
    if (state.view === 'lookup') render()
  })
  document.getElementById('safety')?.addEventListener('change', async (e) => {
    const it = await currentItem()
    if (!it) return
    const n = parseInt(e.target.value, 10)
    it.safety = Number.isFinite(n) && n >= 0 ? n : null
    await putItem(it)
    render()
  })
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
// updateViaCache: 'none'＝檢查新版時不用手機暫存；回到 App 時也檢查一次（新版裝好會自動重新整理）
if ('serviceWorker' in navigator) {
  navigator.serviceWorker
    .register('sw.js', { updateViaCache: 'none' })
    .then((reg) => {
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible' && !state.busy && state.view !== 'analyzing') reg.update().catch(() => {})
      })
    })
    .catch(() => {})
}

/** 設定 →「檢查更新」：問 GitHub 最新版本；不一樣就清掉 App 的暫存（不會動到盤點紀錄、樣品照、API Key）再重新整理 */
async function forceUpdate() {
  let latest = ''
  try {
    const text = await (await fetch(`app.js?check=${Date.now()}`, { cache: 'no-store' })).text()
    latest = (text.match(/const VERSION = '([^']+)'/) || [])[1] || ''
  } catch {
    return toast('連不到 GitHub，請確認有網路')
  }
  if (latest && latest === VERSION) return toast(`已經是最新版：${VERSION}`)
  toast(`更新到 ${latest || '最新版'}…`)
  for (const r of (await navigator.serviceWorker?.getRegistrations?.()) || []) await r.unregister()
  for (const k of await caches.keys()) await caches.delete(k)
  location.replace(`./?v=${Date.now()}`)
}

render()
// 打開 App：有開同步就先跟大家對一次
if (syncReady()) setTimeout(() => syncNow().catch(() => {}), 1200)
