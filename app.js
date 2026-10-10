/**
 * 聖佳智慧庫存（聖佳冷凍材料；原名「拍照盤點」）
 * 拍貨架 → Gemini 視覺模型找出每個商品並框起來 → 原圖對照、＋／－ 修正 → 存在手機；總表一次匯出 Excel、同步 Google 試算表。
 * 品項庫：按「完成」時自動長出來（料號、各位置數量、帳面數、安全庫存）；查型號：拍標籤或打型號 → 解讀、店裡有沒有、替代品。
 * 沒有後端：API Key 只存在這支手機（localStorage），照片只送到 Google Gemini 分析。
 */

import { makeXlsx } from './xlsx.js'
import { decode, normalizeModel, looseKey, canon, modelKey, linksFor } from './rules.js'
import { icon } from './icons.js'

const API = 'https://generativelanguage.googleapis.com/v1beta'
const LS = {
  key: 'inventory:apiKey',
  /** 擁有者有沒有把 AI 金鑰放在雲端（Apps Script）給大家共用：'1'＝有 */
  aiShared: 'inventory:aiShared',
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
  // 盤點規則（跟著共用設定同步）＋ 標準答案考試紀錄（只存這台）
  blind: 'inventory:blind',
  recount: 'inventory:recountOn',
  goldenRuns: 'inventory:goldenRuns',
  // 照片改存內容（4.0.6）：舊資料轉好了沒；有照片在這台不見了，要從雲端拿回來
  photoBuf: 'inventory:photoBuf',
  needRepair: 'inventory:needRepair',
  // 從雲端重新下載全部：更新後自動跑過一次了沒（4.0.7）
  fullPullDone: 'inventory:fullPullDone',
  // 雲端 Apps Script 程式碼的版本（2 以上才會把照片分開存）
  serverVer: 'inventory:serverVer',
  // 電腦、平板的側邊選單收起來了沒
  sideCollapsed: 'inventory:sideCollapsed',
  // 點貨對單（4.7 測試版）：拍單子時框的位置（只在這台）、交叉比對開關（只在這台）、最後一次比對結果、更新後從雲端補拿點貨紀錄了沒
  pickCrop: 'inventory:pickCrop',
  // 遮住的欄位（單價、金額）的位置、第一次的說明看過了沒（都只在這台）
  pickMask: 'inventory:pickMask',
  pickCropHelp: 'inventory:pickCropHelp',
  pickCross: 'inventory:pickCross',
  pickCrossLast: 'inventory:pickCrossLast',
  pickPullDone: 'inventory:pickPullDone',
}
/** 檢視者不能用的動作 */
const EDIT_ACTIONS = new Set(['new', 'analyze', 'add', 'delete-session', 'finish', 'save-catalog', 'reset-catalog', 'add-box', 'del-sample', 'clear-all', 'review-doubts', 'item-add', 'import', 'item-edit', 'item-confirm', 'item-merge', 'item-delete', 'move-in', 'move-out', 'book-set', 'equiv-add', 'read-add', 'loc-add', 'safety-pick', 'recount', 'recount-reason', 'recount-adjust', 'recount-keep', 'golden-run', 'bulk-finish', 'bulk-delete', 'erp-import', 'erp-cats', 'erp-link', 'pk-new', 'pk-add', 'pk-step', 'pk-done', 'pk-confirm', 'pk-qty', 'pk-more', 'pk-assign', 'pk-remove', 'pk-finish', 'pk-discard', 'pk-cancel-edit', 'pk-reopen', 'pk-delete'])
/** 權限（跟 Google 雲端硬碟的「共用」一樣） */
const ROLE_LABEL = { owner: '擁有者', manager: '管理員', editor: '編輯者', viewer: '檢視者' }
const ROLE_DESC = { owner: '全部都可以；不能被移除', manager: '可以盤點、修改，也可以邀請、移除人', editor: '可以盤點、修改', viewer: '只能看（可以下載 Excel）' }
const myRole = () => (ls.get(LS.syncKey) ? ls.get(LS.memberRole) || 'editor' : 'owner')
const canEdit = () => myRole() !== 'viewer'
const canManage = () => ['owner', 'manager'].includes(myRole())
/** 盲盤：盤點的人看不到帳面數、差異（擁有者、管理員看得到）。看得到帳面數，盤點就變成「對答案」，抓不到錯 */
const blindMe = () => ls.get(LS.blind, '1') === '1' && !canManage()
/** 有差異先複盤：請另一個人再數一次，兩次一樣才確定，再選原因、由管理員調整帳面 */
const recountOn = () => ls.get(LS.recount, '1') === '1'
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
const VERSION = '4.7.3（10/10・液態玻璃）'

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
/** action：提示旁邊的按鈕（例如「復原」），{ label, run }；有按鈕的提示留久一點 */
function toast(msg, action) {
  document.querySelector('.toast')?.remove()
  const el = document.createElement('div')
  el.className = 'toast'
  el.setAttribute('role', 'status')
  el.textContent = msg
  if (action) {
    const b = document.createElement('button')
    b.className = 'toast-action'
    b.type = 'button'
    b.textContent = action.label
    b.onclick = () => {
      el.remove()
      action.run()
    }
    el.append(b)
  }
  document.body.appendChild(el)
  clearTimeout(toastTimer)
  // 字多的提示留久一點（最少 5 秒、最多 10 秒），不然還沒看完就不見了；點一下提示可以先關掉
  const len = String(msg).length
  toastTimer = setTimeout(() => el.remove(), Math.max(action ? 6000 : 2800, len > 24 ? Math.min(10000, 5000 + (len - 24) * 60) : 0))
  el.addEventListener('click', (e) => e.target === el && el.remove())
}
// iPhone 的 Safari 要有 touchstart 監聽，按鈕按下去才會變色（:active）
document.addEventListener('touchstart', () => {}, { passive: true })

// ───────────────────────── 存檔（IndexedDB，照片也存在手機） ─────────────────────────
// 照片怎麼存：iPhone 的 Safari 會把存在資料庫裡的照片檔（Blob）弄丟（讀的時候出現「The object can not be found here」，
// 照片變空白、存檔和同步失敗）→ 改存照片的內容（ArrayBuffer），跟資料放在同一筆；讀出來再變回 Blob，其他程式照舊用 Blob。
const LOST = { _lost: true }
const bufOf = (b) => (b.arrayBuffer ? b.arrayBuffer() : new Response(b).arrayBuffer())
async function packBlob(b) {
  if (!(b instanceof Blob)) return b
  try {
    return { _buf: await bufOf(b), _type: b.type || 'image/jpeg' }
  } catch {
    return LOST // 這台的照片檔已經不見了
  }
}
// 還沒從雲端抓下來的照片（同步只拿資料，照片打開那次盤點才抓）
const PENDING = { _pending: true }
const unpackBlob = (v) => (v?._buf instanceof ArrayBuffer ? new Blob([v._buf], { type: v._type }) : v?._lost || v?._pending ? undefined : v)
async function packRecord(store, v) {
  if (store === 'picks') return v // 點貨紀錄：純資料，沒有照片
  if (store === 'sessions')
    return {
      ...v,
      photos: await Promise.all(
        (v.photos || []).map(async ({ lost, pending, ...p }) => ({ ...p, blob: p.blob ? await packBlob(p.blob) : pending ? PENDING : LOST })),
      ),
    }
  if (store === 'items') return v.photo ? { ...v, photo: await packBlob(v.photo) } : v
  return { ...v, blob: await packBlob(v.blob) }
}
/** 讀出來：照片變回 Blob；沒有照片的：pending＝還沒下載（打開會抓）、lost＝這台弄丟了（畫面顯示「從雲端拿回」） */
function unpackRecord(store, v) {
  if (!v || store === 'picks') return v
  if (store === 'sessions')
    return {
      ...v,
      photos: (v.photos || []).map((p) => {
        const blob = unpackBlob(p.blob)
        if (blob) return { ...p, blob }
        return p.blob?._pending ? { ...p, blob: undefined, pending: true } : { ...p, blob: undefined, lost: true }
      }),
    }
  if (store === 'items') return v.photo ? { ...v, photo: unpackBlob(v.photo) } : v
  return { ...v, blob: unpackBlob(v.blob) }
}
const hasOldBlob = (store, v) => (store === 'sessions' ? (v.photos || []).some((p) => p.blob instanceof Blob) : (store === 'items' ? v.photo : v.blob) instanceof Blob)
const lostCount = (store, v) => (store === 'sessions' ? (v.photos || []).filter((p) => p.blob?._lost).length : (store === 'items' ? v.photo : v.blob)?._lost ? 1 : 0)

/** sessions＝盤點紀錄；samples＝樣品照（第 2 版新增）；items＝品項庫（第 3 版新增）；erp＝正航產品表（第 4 版）；picks＝點貨紀錄（第 5 版） */
const DB_VERSION = 5
const idb = (() => {
  let p
  /** 舊資料（照片存成 Blob）一筆一筆轉成新存法；讀不到的照片記成「不見了」，同步時從雲端拿回來 */
  const migrate = async (d) => {
    if (ls.get(LS.photoBuf) === '1') return
    const run = (store, mode, fn) =>
      new Promise((resolve, reject) => {
        const t = d.transaction(store, mode)
        const r = fn(t.objectStore(store))
        t.oncomplete = () => resolve(r?.result)
        t.onerror = () => reject(t.error)
        t.onabort = () => reject(t.error)
      })
    let lost = 0
    try {
      for (const store of ['sessions', 'items', 'samples']) {
        for (const k of await run(store, 'readonly', (s) => s.getAllKeys())) {
          const v = await run(store, 'readonly', (s) => s.get(k))
          if (!v || !hasOldBlob(store, v)) continue
          const packed = await packRecord(store, v)
          lost += lostCount(store, packed)
          await run(store, 'readwrite', (s) => s.put(packed))
        }
      }
      ls.set(LS.photoBuf, '1')
    } catch (e) {
      console.error('migrate', e)
    }
    if (lost) ls.set(LS.needRepair, '1')
  }
  const open = () => (p ??= openRaw().then(async (d) => (await migrate(d), d)))
  // withVersion＝false：這台的資料庫比這版 App 新（例如以後的新版用過、又退回這一版）→ 不帶版本號照現況打開，
  // 不升級、不刪資料（少的那一格用到時才會出錯，至少盤點紀錄、品項庫照常可以用）
  const openRaw = (withVersion = true) =>
    new Promise((resolve, reject) => {
      const req = withVersion ? indexedDB.open('inventory', DB_VERSION) : indexedDB.open('inventory')
      req.onupgradeneeded = () => {
        const d = req.result
        if (!d.objectStoreNames.contains('sessions')) d.createObjectStore('sessions', { keyPath: 'id' })
        if (!d.objectStoreNames.contains('erp')) d.createObjectStore('erp', { keyPath: 'id' }) // 第 4 版：正航產品表（一筆就是整張表）
        if (!d.objectStoreNames.contains('picks')) d.createObjectStore('picks', { keyPath: 'id' }) // 第 5 版：點貨紀錄（只加新的一格，舊資料不動）
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
      req.onerror = () => {
        if (withVersion && req.error?.name === 'VersionError') return openRaw(false).then(resolve, reject)
        reject(req.error)
      }
    })
  const tx = async (store, mode, fn, retry = true) => {
    const d = await open()
    try {
      return await new Promise((resolve, reject) => {
        const t = d.transaction(store, mode)
        const r = fn(t.objectStore(store))
        t.oncomplete = () => resolve(r?.result)
        t.onerror = () => reject(t.error)
        t.onabort = () => reject(t.error || new DOMException('存檔被中斷', 'AbortError'))
      })
    } catch (e) {
      // iPhone 放到背景一陣子，資料庫連線可能被系統收掉（UnknownError／InvalidStateError）：重新連一次再試
      if (retry && ['InvalidStateError', 'UnknownError', 'AbortError'].includes(e?.name)) {
        try {
          d.close()
        } catch {}
        p = undefined
        return tx(store, mode, fn, false)
      }
      throw e
    }
  }
  const storeOf = (store) => ({
    // 樣品照的照片不見了就不拿出來用（AI 比對、上傳都會出錯）
    all: async () =>
      ((await tx(store, 'readonly', (s) => s.getAll())) ?? [])
        .map((v) => unpackRecord(store, v))
        .filter((v) => store !== 'samples' || v.blob)
        .sort((a, b) => b.createdAt - a.createdAt),
    get: async (id) => unpackRecord(store, await tx(store, 'readonly', (s) => s.get(id))),
    /** 一般存檔：記下修改時間（多台同步靠它判斷哪一份比較新），稍後自動同步 */
    put: async (item) => {
      item.updatedAt = Date.now()
      scheduleSync()
      const rec = await packRecord(store, item)
      return tx(store, 'readwrite', (s) => s.put(rec))
    },
    /** 原樣存（同步下載的、只改本機狀態的）：不改修改時間、不觸發同步 */
    putRaw: async (item) => {
      const rec = await packRecord(store, item)
      return tx(store, 'readwrite', (s) => s.put(rec))
    },
    /** 上傳成功：如果這段時間沒再改過，標記「已同步」（同一個交易裡讀和寫，不會蓋掉剛改的） */
    markSynced: (id, t) =>
      tx(store, 'readwrite', (s) => {
        const req = s.get(id)
        req.onsuccess = () => {
          const v = req.result
          // 還是舊存法（照片是 Blob）的不要原樣再存一次：iPhone 會因此把照片檔弄丟
          if (v && (v.updatedAt || v.createdAt) === t && !hasOldBlob(store, v)) s.put({ ...v, _syncT: t })
        }
        return req
      }),
    /** 照片傳上雲端了：在存的那筆上記 up＝true（不改修改時間、不觸發同步） */
    markPhotoUp: (id, pids) =>
      tx(store, 'readwrite', (s) => {
        const req = s.get(id)
        req.onsuccess = () => {
          const v = req.result
          if (!v?.photos) return
          let hit = false
          for (const p of v.photos) if (pids.includes(p.id)) hit = p.up = true
          if (hit) s.put(v)
        }
        return req
      }),
    del: (id) => tx(store, 'readwrite', (s) => s.delete(id)),
    clear: () => tx(store, 'readwrite', (s) => s.clear()),
  })
  return { sessions: storeOf('sessions'), samples: storeOf('samples'), items: storeOf('items'), erp: storeOf('erp'), picks: storeOf('picks') }
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
/** 把照片讀成可以畫的圖：有些 iPhone 的 createImageBitmap 會失敗，改用一般的 <img> 讀 */
async function bitmapOf(blob) {
  try {
    return await createImageBitmap(blob)
  } catch {
    const img = new Image()
    const url = URL.createObjectURL(blob)
    img.src = url
    try {
      await img.decode()
    } finally {
      URL.revokeObjectURL(url)
    }
    return Object.assign(img, { width: img.naturalWidth, height: img.naturalHeight })
  }
}
async function cropBox(photo, box) {
  const bmp = await bitmapOf(photo.blob)
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
/** 照片網址：同一張（大小一樣）沿用，不要每次重畫都重新載入；照片不見了回傳空字串 */
const urlOf = (photo) => {
  if (!photo.blob) return ''
  const hit = urls.get(photo.id)
  if (hit && (hit.blob === photo.blob || hit.size === photo.blob.size)) return hit.url
  if (hit) URL.revokeObjectURL(hit.url)
  const url = URL.createObjectURL(photo.blob)
  urls.set(photo.id, { blob: photo.blob, size: photo.blob.size, url })
  return url
}
/** 照片不用了（刪掉、復原時間過了、整次盤點刪掉）：放掉它占的記憶體 */
const dropUrl = (id) => {
  const hit = urls.get(id)
  if (!hit) return
  URL.revokeObjectURL(hit.url)
  urls.delete(id)
}
/**
 * 框的放大圖：用 SVG 的 viewBox 只露出那一塊（不切圖、不用 canvas，iPhone 也穩），框線用那一種的顏色。
 * ratio＝放大圖格子的寬÷高：把露出的範圍撐到一樣比例、而且不超出照片，格子才不會有一邊空白。
 */
function boxZoomSvg(photo, box, color = '#0a84ff', ratio = 0) {
  const W = photo.w || 1000
  const H = photo.h || 1000
  const [a, b, c, d] = box
  const x1 = (Math.min(b, d) / 1000) * W
  const x2 = (Math.max(b, d) / 1000) * W
  const y1 = (Math.min(a, c) / 1000) * H
  const y2 = (Math.max(a, c) / 1000) * H
  const pad = Math.max(x2 - x1, y2 - y1) * 0.15 + 8
  let vx = Math.max(0, x1 - pad)
  let vy = Math.max(0, y1 - pad)
  let vw = Math.min(W, x2 + pad) - vx
  let vh = Math.min(H, y2 + pad) - vy
  if (ratio > 0 && Number.isFinite(ratio)) {
    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v))
    if (vw / vh < ratio) {
      const nw = Math.min(W, vh * ratio)
      vx = clamp(vx + vw / 2 - nw / 2, 0, W - nw)
      vw = nw
    } else {
      const nh = Math.min(H, vw / ratio)
      vy = clamp(vy + vh / 2 - nh / 2, 0, H - nh)
      vh = nh
    }
  }
  return `<svg viewBox="${vx.toFixed(1)} ${vy.toFixed(1)} ${vw.toFixed(1)} ${vh.toFixed(1)}" preserveAspectRatio="xMidYMid meet" role="img" aria-label="這一個框的放大圖">
    <image href="${urlOf(photo)}" x="0" y="0" width="${W}" height="${H}" preserveAspectRatio="none"/>
    <path d="M0 0H${W}V${H}H0Z M${x1.toFixed(1)} ${y1.toFixed(1)}V${y2.toFixed(1)}H${x2.toFixed(1)}V${y1.toFixed(1)}Z" fill="#000" fill-opacity="0.35" fill-rule="evenodd"/>
    <rect x="${x1.toFixed(1)}" y="${y1.toFixed(1)}" width="${(x2 - x1).toFixed(1)}" height="${(y2 - y1).toFixed(1)}" fill="none" stroke="${color}" stroke-width="3" vector-effect="non-scaling-stroke" rx="2"/>
  </svg>`
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
  // 這台沒有自己的金鑰、擁有者有放共用金鑰：改由自己的 Apps Script 去問 Gemini（金鑰不會到這支手機）
  if (!key && aiViaSheet()) return callViaSheet(path, opts, timeoutMs)
  let res
  const ctrl = new AbortController()
  const timer = timeoutMs ? setTimeout(() => ctrl.abort(), timeoutMs) : 0
  // opts.signal：另一個模型已經先回答了，這邊就取消
  const outer = opts.signal
  const onOuter = () => ctrl.abort()
  outer?.addEventListener('abort', onOuter)
  try {
    if (outer?.aborted) throw new Error('aborted')
    // 金鑰放在標頭（x-goog-api-key），不放網址：網址可能被記錄在紀錄檔或截圖裡
    res = await fetch(`${API}/${path}`, { ...opts, headers: { ...(opts.headers || {}), 'x-goog-api-key': key }, signal: ctrl.signal })
  } catch {
    if (outer?.aborted) throw new ApiError('已取消', 499, 'cancelled')
    if (ctrl.signal.aborted) throw new ApiError(`等了 ${Math.round(timeoutMs / 1000)} 秒還沒回，自動重試。`, 408, `timeout ${timeoutMs / 1000}s`)
    throw new ApiError('沒有網路，或連不到 Google。辨識時請不要切到別的 App、不要讓螢幕暗掉。', 0, 'network error')
  } finally {
    clearTimeout(timer)
    outer?.removeEventListener('abort', onOuter)
  }
  const data = await res.json().catch(() => ({}))
  return toResult(res.status, data)
}
/** Gemini 的回覆：成功就回資料；失敗換成看得懂的錯誤（直接連、經過 Apps Script 都用這個） */
function toResult(status, data) {
  // 沒有狀態碼（舊版雲端程式碼、回覆壞掉）也算失敗，不能當成功
  if (!(status >= 200 && status < 300)) {
    const msg = data?.error?.message || ''
    const err = new ApiError(friendly(status, msg), status, `${status} ${data?.error?.status || ''} ${msg}`.trim().slice(0, 200))
    // 429 會告訴你要等多久（例如 "17s"）
    const retry = (data?.error?.details || []).find((d) => d.retryDelay)?.retryDelay
    if (retry) err.retryAfter = Math.min(30, parseFloat(retry) || 0)
    throw err
  }
  return data
}
/** 有沒有 AI 可以用：自己的金鑰，或擁有者放在雲端的共用金鑰 */
function hasAi() {
  return !!ls.get(LS.key) || aiViaSheet()
}
/** 擁有者有把金鑰放在雲端共用（不管這台能不能用） */
function aiSharedOn() {
  return !!(ls.get(LS.sheet) && ls.get(LS.syncKey)) && ls.get(LS.aiShared) === '1'
}
/** 這台可以用共用金鑰：檢視者不行（雲端也會擋），所以檢視者的拍照鈕不會亮 */
function aiViaSheet() {
  return aiSharedOn() && canEdit()
}
/** 經過 Apps Script 時多等的時間：雲端那邊收到後照樣會跑完、照樣計費，太早放棄再送一次＝花兩次錢 */
const SHEET_GRACE = 20000
/** 共用金鑰：請求送到自己的 Apps Script，由它拿雲端的金鑰問 Gemini（會多 1～2 秒） */
async function callViaSheet(path, opts, timeoutMs) {
  const outer = opts.signal
  if (outer?.aborted) throw new ApiError('已取消', 499, 'cancelled')
  const ctrl = new AbortController()
  let timedOut = false
  const timer = timeoutMs
    ? setTimeout(() => {
        timedOut = true
        ctrl.abort()
      }, timeoutMs + SHEET_GRACE)
    : 0
  const onAbort = () => ctrl.abort()
  outer?.addEventListener('abort', onAbort)
  let r
  try {
    r = await postSync({ action: 'ai', path, body: opts.body || '' }, { signal: ctrl.signal })
  } catch (e) {
    if (outer?.aborted) throw new ApiError('已取消', 499, 'cancelled')
    if (timedOut) throw new ApiError(`等了 ${Math.round((timeoutMs + SHEET_GRACE) / 1000)} 秒還沒回，自動重試。`, 408, `timeout ${(timeoutMs + SHEET_GRACE) / 1000}s (sheet)`)
    throw sheetAiError(e)
  } finally {
    clearTimeout(timer)
    outer?.removeEventListener('abort', onAbort)
  }
  // 舊版雲端程式碼不認得 ai：回覆裡沒有狀態碼
  if (typeof r.status !== 'number') throw new ApiError('雲端的程式碼是舊版，還不會幫忙問 AI：請擁有者到「設定 → Google 試算表」複製新的程式碼、貼上、部署新版本。', 403, 'sheet: no status (old Apps Script)')
  let data = {}
  try {
    data = JSON.parse(r.body || '{}')
  } catch {
    /* 回覆不是 JSON：當成空的，toResult 會依狀態碼處理 */
  }
  return toResult(r.status, data)
}
/**
 * 經過雲端問 AI 失敗：分清楚是「雲端說不行」（檢視者、被移除、今天次數到上限、不支援 → 不重試，直接顯示原因）
 * 還是「網路斷了、雲端暫時出錯」（可以重試）。以前全部當成網路斷線，一張照片會自動重試十幾次。
 */
function sheetAiError(e) {
  const d = e?.data
  const msg = e?.message || '連不到 Google（沒有網路？）'
  if (!d || d.network) return new ApiError(msg, 0, `sheet ${msg}`.slice(0, 200))
  if (d.html) return Object.assign(new ApiError(msg, 503, `sheet html ${d.httpStatus || ''}`.trim()), { sheetDown: true })
  // 404＝共用金鑰不支援這個模型：換下一個模型試；其他都是雲端明確說不行：重試也一樣，不要浪費
  const status = d.status === 404 && !d.quota && !d.viewer && !d.revoked ? 404 : 403
  return new ApiError(d.error || msg, status, `sheet ${status} ${d.error || ''}`.trim().slice(0, 200))
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
  // 儲位代號限 20 字：AI 偶爾把整段標籤文字塞進來，會撐爆畫面
  return { objects: kept, note, model, location: String(parsed.location || '').trim().slice(0, 20) }
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
  // 用共用金鑰（經過 Apps Script）不要「同時請備用模型」：這邊取消了，雲端那邊照樣跑完、照樣計費，等於花兩次錢
  if (!backup || viaSheetNow()) return primary.finally(() => hedge.abort())
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

/** 這次的 AI 請求會經過 Apps Script（這台沒有自己的金鑰、用擁有者的共用金鑰） */
const viaSheetNow = () => !ls.get(LS.key) && aiViaSheet()
/** 連續幾次「網路斷了」就停：網路真的不通，再試也是白等 */
const MAX_NET_FAILS = 3

/** 依序試：同一個模型（有思考設定 → 拿掉 → 簡化），不行就換下一個模型 */
async function attemptChain(models, image, refs, onStatus, signal) {
  let lastErr
  let netFails = 0
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
          // 共用金鑰：雲端那邊可能還在跑（會計費），先等一下再送下一個，不要兩個同時跑
          if (viaSheetNow() && mi < models.length - 1) {
            onStatus(`${model} 太慢，等 3 秒再換比較快的模型`)
            await sleep(3000)
          } else onStatus(`${model} 太慢，換比較快的模型`)
          break
        }
        // 400：先拿掉思考設定、再改簡化請求；都不行就是別的問題（Key、照片）
        if (e.status === 400) {
          if (pi < plan.length - 1) continue
          throw e
        }
        if (!retryable(e)) throw e
        // 網路斷了、Apps Script 整個出錯（回網頁不是 JSON）：連續 3 次就停，換模型也沒用
        if (e.status === 0 || e.sheetDown) {
          if (++netFails >= MAX_NET_FAILS) throw e
          onStatus(e.sheetDown ? 'Google 試算表那邊出錯，再試一次' : '網路斷了一下，重新送出')
        } else netFails = 0
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
/** 畫面用：品牌・型號・規格，每一段不拆開換行（「3分」的「分」不會自己掉到下一行） */
const detailHtml = (o) => [o.brand, o.model, o.spec].filter(Boolean).map((v) => `<span class="nb">${esc(v)}</span>`).join('・')

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
  session.photos.forEach((photo, pi) => photoDoubts(photo, (o) => o.checked || o.edited).forEach((d) => out.push({ pi, ...d })))
  return out
}
/** 一張照片裡要確認的框（skip＝不用再列的，例如你確認過、改過的） */
function photoDoubts(photo, skip = () => false) {
  const out = []
  if (!photo.w || !photo.h) return out
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
      if (skip(o)) return
      const diff = sizes[i] / med - 1
      let reason = ''
      if (o.odd) reason = o.oddReason ? `AI 說這一個${o.oddReason}` : 'AI 覺得跟同一種的其他個不太一樣'
      else if (objs.length >= 3 && Math.abs(diff) > 0.3) reason = `比同一種的其他個${diff > 0 ? '大' : '小'}約 ${Math.round(Math.abs(diff) * 100)}%`
      else if (pairGap > 1.35) reason = '這一種只有 2 個，但大小差很多'
      else if (o.confidence < 0.6) reason = 'AI 不太確定這是什麼'
      if (reason) out.push({ o, reason })
    })
  }
  return out
}

// ───────────────────────── AI 準不準：記下 AI 原本的答案，人工修改後就知道對不對 ─────────────────────────
/**
 * 辨識（和背景的相似品比對）一好，就把 AI 的答案記在每個框上：o.ai＝AI 說的種類、o.aiFlag＝AI 有沒有標成「要確認」。
 * 之後你補框（o.added）、刪框（記在 photo.aiGone）、改種類、改數量，按「完成」時就能算出 AI 原本對不對。
 */
function markAi(photo, objs = photo.objects) {
  const flagged = new Set(photoDoubts(photo).map((d) => d.o))
  for (const o of objs) {
    if (o.added) continue
    o.ai = keyOf(o)
    o.aiFlag = flagged.has(o)
  }
}

/**
 * 這次盤點 AI 準不準（按「完成」時算，存在 session.eval）：
 * - 每一種：AI 數幾個 vs 你確認後幾個。AI 的一種，對到「它的框最後大多變成的那一種」，所以整組改名不算錯，數量對就算對。
 * - 拆開看：漏數（你補的框＋你把數量加上去的）、多數（刪掉的框＋你把數量減下來的）、分錯（框改成別的種類）。
 * - 要確認抓到幾個：AI 錯的框（刪掉、分錯）裡，有幾個事先被標成「要確認」。越高，代表只看要確認的就夠。
 * 手動新增的種類（沒有框）不算；舊版本的盤點沒有記 AI 答案，回傳 null。
 */
function evalSession(s) {
  const done = s.photos.filter((p) => p.status === 'done')
  const objs = done.flatMap((p) => p.objects)
  const gone = done.flatMap((p) => p.aiGone || [])
  if (!objs.some((o) => o.ai) && !gone.length) return null
  const bump = (m, k, n = 1) => m.set(k, (m.get(k) || 0) + n)
  const aiCount = new Map()
  for (const o of objs) if (o.ai) bump(aiCount, o.ai)
  for (const x of gone) bump(aiCount, x.ai)
  const votes = new Map()
  for (const o of objs) {
    if (!o.ai) continue
    if (!votes.has(o.ai)) votes.set(o.ai, new Map())
    bump(votes.get(o.ai), keyOf(o))
  }
  const mapTo = new Map([...votes].map(([a, v]) => [a, [...v].sort((x, y) => y[1] - x[1])[0][0]]))
  const aiOf = new Map()
  const types = []
  for (const [a, n] of aiCount) {
    const f = mapTo.get(a)
    if (f) bump(aiOf, f, n)
    else {
      const [label, , , spec] = a.split('|')
      types.push({ label, spec, ai: n, final: 0, item: '' }) // AI 說有、其實沒有的一種（框全被刪掉）
    }
  }
  const groups = groupsOf(s).filter((g) => !g.manual)
  // 還有「要確認」沒看過的那一種：不知道 AI 對不對，不算進準確率（以前會被當成「一個不差」）
  const open = new Set(doubtsOf(s).map((d) => keyOf(d.o)))
  for (const g of groups) if (!open.has(g.key)) types.unshift({ label: g.label, spec: g.spec, ai: aiOf.get(g.key) || 0, final: Number(g.count) || 0, item: s.itemOf?.[g.key] || '' })
  let missed = objs.filter((o) => o.added).length
  let extra = gone.length
  for (const g of groups) {
    const d = (Number(g.count) || 0) - g.boxes
    if (d > 0) missed += d
    else extra -= d
  }
  const wrongObjs = objs.filter((o) => o.ai && keyOf(o) !== mapTo.get(o.ai))
  const wrong = wrongObjs.length + gone.length
  const caught = wrongObjs.filter((o) => o.aiFlag).length + gone.filter((x) => x.flag).length
  return {
    at: Date.now(),
    model: s.model || '',
    boxes: objs.filter((o) => o.ai).length + gone.length,
    types,
    missed,
    extra,
    misclass: wrongObjs.length,
    renamed: [...mapTo].filter(([a, f]) => a !== f).length,
    wrong,
    caught,
    flagged: objs.filter((o) => o.aiFlag).length + gone.filter((x) => x.flag).length,
    unchecked: groups.filter((g) => open.has(g.key)).length,
  }
}

/** 很多次盤點合起來算：完全正確率、平均差幾個、漏數／多數／分錯、要確認抓到幾成 */
function qualityStats(sessions) {
  const types = sessions.flatMap((s) => s.eval.types)
  const n = types.length
  const exact = types.filter((t) => t.ai === t.final).length
  const sum = (k) => sessions.reduce((a, s) => a + (s.eval[k] || 0), 0)
  return {
    sessions: sessions.length,
    n,
    exact,
    rate: n ? exact / n : null,
    mae: n ? types.reduce((a, t) => a + Math.abs(t.ai - t.final), 0) / n : null,
    boxes: sum('boxes'),
    missed: sum('missed'),
    extra: sum('extra'),
    misclass: sum('misclass'),
    wrong: sum('wrong'),
    caught: sum('caught'),
    flagged: sum('flagged'),
    unchecked: sum('unchecked'),
  }
}
const pct = (x) => (x == null ? '—' : `${Math.round(x * 100)}%`)

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
  // AI 框被刪掉＝AI 多數了（記下來算準確率）
  for (const r of refs) {
    const photo = session.photos[r.pi]
    const o = photo.objects[r.oi]
    if (o?.ai) (photo.aiGone ??= []).push({ ai: o.ai, flag: !!o.aiFlag })
  }
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
/**
 * 找儲位：品項清單每一格位置都要找一次（3000 個品項就是幾千次），不能每次都重新讀設定、解析 JSON。
 * 用「代號 → 儲位」的對照表；設定的內容（原始字串）變了才重建。
 */
let locIndex = { raw: null, map: new Map() }
const findLocation = (code) => {
  const raw = ls.get(LS.locations, '[]')
  if (raw !== locIndex.raw) {
    const map = new Map()
    for (const l of locations()) {
      const k = canon(l?.code)
      if (k && !map.has(k)) map.set(k, l) // 同一個代號建了兩次：照舊用第一個
    }
    locIndex = { raw, map }
  }
  return locIndex.map.get(canon(code))
}
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
/**
 * 畫面上的儲位一律畫成「店裡貼的黃底黑字標籤」（4.6 貨架標籤外觀）：A-03＋說明。
 * size：'sm'（清單裡）、'lg'（儲位頁）、'xl'（盤點結果的標題）；沒填位置＝灰色虛線框。
 */
const locTag = (code, size = '') => `<span class="loc-tag${size ? ` ${size}` : ''}">${esc(code)}</span>`
const placeHtml = (place, { size = '', name = true, none = '沒填位置' } = {}) => {
  if (!place) return `<span class="loc-tag none${size ? ` ${size}` : ''}">${esc(none)}</span>`
  const loc = findLocation(place)
  return `<span class="loc">${locTag(loc ? loc.code : place, size)}${name && loc?.name ? `<span class="loc-name">${esc(loc.name)}</span>` : ''}</span>`
}
/** 一串「位置 數量」：A-01 ×3　A-02 ×2（品項清單、產品總表、總表用） */
const placesHtml = (pairs, max = 99) => {
  const shown = pairs.slice(0, max)
  return `<span class="loc-list">${shown.map(([p, n]) => `<span class="loc-pair">${p && !String(p).startsWith('未填位置') ? locTag(findLocation(p)?.code || p, 'sm') : `<span class="loc-tag none sm">${esc(p || '沒填位置')}</span>`}${n != null ? `<span class="loc-n">×${esc(n)}</span>` : ''}</span>`).join('')}${pairs.length > shown.length ? `<span class="loc-n">等 ${pairs.length} 格</span>` : ''}</span>`
}

// ───────────────────────── 品項庫：盤點時自動長出來 ─────────────────────────
/**
 * 每一種商品一筆，料號 P0001 起。按「完成」時，這次盤點的每一種都會對到品項庫：
 * 對得到 → 記下這個位置的數量；對不到 → 自動建立「新的」品項（之後在「品項」確認名稱、合併重複的）。
 * 實盤＝每個位置「最近一次」數到的數量加起來；帳面數＝應該要有幾個（進貨加、賣出減，或拿實盤當起點）。
 */
let itemsCache = null
/** 品項庫改過幾次（改名、接上正航、新增、刪除）：點貨的搜尋對照表看這個決定要不要重建 */
let itemsVer = 0
/** 很舊的品項（或別台舊版傳來的）可能少了這些欄位：讀進來時補上，後面的程式才不會出錯（例如匯入正航時 moves.push） */
const fixItem = (it) => {
  if (!Array.isArray(it.moves)) it.moves = []
  if (!it.stock || typeof it.stock !== 'object') it.stock = {}
  if (!Array.isArray(it.aliases)) it.aliases = []
  if (!Array.isArray(it.equiv)) it.equiv = []
  return it
}
async function itemsAll(force = false) {
  if (!itemsCache || force) {
    itemsCache = (await idb.items.all().catch(() => [])).map(fixItem)
    itemsVer++
  }
  return itemsCache
}
async function putItem(it) {
  it.updatedAt = Date.now()
  await idb.items.put(fixItem(it))
  if (itemsCache && !itemsCache.includes(it)) itemsCache.unshift(it)
  itemsVer++
}
async function delItem(it) {
  await idb.items.del(it.id)
  tombstone(`item:${it.id}`)
  if (itemsCache) itemsCache = itemsCache.filter((x) => x !== it)
  itemsVer++
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
/**
 * 叫貨看的數字（待使用者／老闆確認規則）：實盤、帳面兩個都有就取比較小的——寧可多提醒，
 * 例如帳面 8、架上只剩 3，也要提醒叫貨。沒盤過就看帳面；沒有帳面就看實盤。
 */
const orderQty = (it) => {
  const counted = liveStock(it).length ? onHand(it) : null
  if (it.book == null) return counted ?? 0
  return counted == null ? it.book : Math.min(it.book, counted)
}
/** 畫面上說明用哪個數字（盲盤的人看不到帳面數，不寫數字） */
const orderWhy = (it) => {
  if (blindMe()) return '數量已經到叫貨點'
  const counted = liveStock(it).length ? onHand(it) : null
  if (it.book != null && counted != null) return `現在剩 ${orderQty(it)} 個（實盤 ${counted}、帳面 ${it.book}，看比較少的）`
  return `現在剩 ${orderQty(it)} 個（${it.book != null ? '帳面' : '實盤'}）`
}
const needsOrder = (it) => it.safety != null && orderQty(it) <= it.safety
const placeKeyOf = (s) => (s.place ? `p:${canon(s.place)}` : `s:${s.id}`)

/**
 * 複盤：按「完成」後，這次數到的跟帳面不一樣 → 開一張複盤單（it.recounts[盤點ID]）。
 * 另一個人再數一次（看不到第一次的數字）：兩次一樣＝差異確定 → 選原因 → 管理員決定要不要調整帳面；
 * 不一樣＝以新數字為準、請再數一次；改完跟帳面一樣＝第一次數錯，自動更正。
 */
const REASONS = ['數錯了（第一次數錯）', '放錯格', '漏記進貨', '漏記賣出、出貨', '損壞、報廢', '借出、當樣品', '單位、包裝算錯', '其他']
const RC_TEXT = { pending: '待複盤', confirmed: '差異確定', fixed: '數錯，已更正', adjusted: '已調整帳面', kept: '保留，不調整' }
const rcOpen = (rc) => !!rc && !rc.gone && (rc.status === 'pending' || rc.status === 'confirmed')
const recountsOf = (it) =>
  Object.entries(it.recounts || {})
    .filter(([, rc]) => !rc.gone)
    .sort((a, b) => rcOpen(b[1]) - rcOpen(a[1]) || b[1].at - a[1].at)
const needsRecount = (it) => recountsOf(it).some(([, rc]) => rcOpen(rc))
/** 試算表、Excel 用：最需要處理的那一張（沒有就是最近的） */
const latestRecount = (it) => recountsOf(it)[0]?.[1] || null
function planRecount(it, s) {
  const pk = placeKeyOf(s)
  const st = it.stock?.[pk]
  const rcs = (it.recounts ??= {})
  const cur = rcs[s.id]
  const now = Date.now()
  // 同一格之前沒處理完的單：被這次盤點取代
  for (const [sid, rc] of Object.entries(rcs)) if (sid !== s.id && rc.pk === pk && rcOpen(rc)) rcs[sid] = { ...rc, gone: true, v: now }
  const d = diffOf(it)
  if (!st || st.removed || st.sid !== s.id || !d) {
    // 改完再按完成，已經沒有差異：還沒複盤的單拿掉
    if (cur && cur.status === 'pending' && !cur.gone) rcs[s.id] = { ...cur, gone: true, v: now }
    return
  }
  if (cur && !cur.gone && cur.first.count === st.count) return // 重按完成、數字沒變：照舊
  rcs[s.id] = { pk, place: s.place || '', at: s.createdAt, first: { count: st.count, by: byName(s), byId: s.byId || '' }, counts: [], status: 'pending', v: now }
}
/** 兩台同時改複盤單：每一張比寫入時間，新的贏 */
function mergeRecounts(a = {}, b = {}) {
  const out = {}
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) out[k] = (a[k]?.v || 0) >= (b[k]?.v || 0) ? a[k] || b[k] : b[k]
  return out
}
const stockPlace = (st) => (st.place ? placeLabel(st.place) : `沒填位置・${fmtTime(st.at)}`)
const itemTitle = (it) => `${it.label}${it.spec ? `・${it.spec}` : ''}`
const itemUrls = new Map()
const itemUrl = (it) => {
  if (!it.photo) return ''
  const hit = itemUrls.get(it.id)
  if (hit && (hit.blob === it.photo || hit.size === it.photo.size)) return hit.url
  if (hit) URL.revokeObjectURL(hit.url)
  const url = URL.createObjectURL(it.photo)
  itemUrls.set(it.id, { blob: it.photo, size: it.photo.size, url })
  return url
}
/** 品項放大看時下面的說明：品名・料號，加上放在哪幾格（黃標籤） */
const itemZoom = (it, inRow = false) => zoomAttrs(itemUrl(it), [itemTitle(it), it.no].filter(Boolean).join('・'), { places: [...new Set(liveStock(it).map(([, st]) => st.place).filter(Boolean))].slice(0, 6), inRow })
/** 品項縮圖：有照片的點一下就放大看（不會觸發那一列原本的點擊）；沒照片的灰色字母方塊不用。inRow：包在 <button class="row"> 裡 */
const itemThumb = (it, size = 44, inRow = false) =>
  it.photo
    ? zoomWrap(`<img class="thumb" src="${itemUrl(it)}" alt="" loading="lazy" decoding="async" style="width:${size}px;height:${size}px">`, itemZoom(it, inRow), size >= 56 ? 'md' : '')
    : `<span class="thumb ph" style="width:${size}px;height:${size}px" aria-hidden="true">${esc(it.label.slice(0, 1))}</span>`

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
    if (st && !st.removed && !sums.has(it) && ((st.sid === s.id && !st.rc) || st.at <= s.createdAt)) {
      it.stock[pk] = removedEntry(st, s.createdAt)
      await putItem(it)
    }
  }
  for (const [it, count] of sums) {
    const st = it.stock[pk]
    if (!st || (st.sid === s.id && !st.rc) || st.at <= s.createdAt) it.stock[pk] = { count, at: s.createdAt, sid: s.id, place: s.place || '', v: Date.now() }
    if (recountOn() && !it.stock[pk]?.rc) planRecount(it, s)
    await putItem(it)
  }
  s.linkedAt = Date.now()
  s.eval = evalSession(s) // AI 準不準（AI 原本的答案 vs 你確認後的）
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

/** 新盤點刪掉、還可以「復原」的照片（只放在記憶體，不存檔） */
let removedPhotos = { s: null, list: [], timer: 0 }
/** 新版 App 已經下載好、等安全的時候重新整理（見最下面 tryReload）；updateTold＝提示過了 */
let updateReady = false
let updateTold = false

function go(view, extra = {}) {
  // 離開首頁就結束多選
  if (view !== 'home') state.selecting = false
  Object.assign(state, { view, focus: null, focusObj: null, addMode: false, viewer: false }, extra)
  render()
  window.scrollTo({ top: 0 })
}
/** 按了要等一下的按鈕：先停用、改成「…中」，做完還原（畫面重畫的話就自然換掉了） */
function busyBtn(btn, text) {
  if (!btn) return () => {}
  const html = btn.innerHTML
  btn.disabled = true
  btn.setAttribute('aria-busy', 'true')
  btn.textContent = text
  return () => {
    if (!btn.isConnected) return
    btn.disabled = false
    btn.removeAttribute('aria-busy')
    btn.innerHTML = html
  }
}
/** 存檔失敗一定要讓人看到原因（不然點了像沒反應） */
function saveFailed(e) {
  const quota = e?.name === 'QuotaExceededError'
  toast(quota ? '手機空間不夠，存不進去：請刪掉一些舊的盤點紀錄或手機裡的照片' : `沒有存進手機（${e?.name || e?.message || '不明原因'}）：請把 App 完全關掉（往上滑掉）再打開，再試一次`)
}
async function save() {
  const s = state.session
  if (!s) return
  try {
    await db.put(s)
  } catch (e) {
    // iPhone 切回 App 後，之前讀出來的照片有時候讀不到了 → 用資料庫裡的照片重新組一份再存一次
    try {
      const fresh = await db.get(s.id)
      for (const ph of s.photos) {
        const f = fresh?.photos?.find((x) => x.id === ph.id)
        if (f?.blob) {
          ph.blob = f.blob
          delete ph.lost
        }
      }
      await db.put(s)
    } catch (e2) {
      console.error('save', e, e2)
      saveFailed(e2)
      throw e2
    }
  }
}

// ───────────────────────── 畫面 ─────────────────────────
const chev = '<svg class="chev" width="10" height="17" viewBox="0 0 10 17" aria-hidden="true"><path d="M1.5 1.5 8 8.5l-6.5 7" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>'
const backBtn = (to = 'home', label = '盤點') => `<button class="back" data-go="${to}"><svg width="12" height="20" viewBox="0 0 12 20" aria-hidden="true"><path d="M10 2 2 10l8 8" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>${esc(label)}</button>`
/** 底部分頁列（像 iOS 的 Tab Bar）：盤點／品項／查型號（圖示用 icons.js 同一套線條） */
// 設定：手機在首頁右上角；平板、電腦放進右邊側邊欄（wide）
// 點貨（4.7 測試版）放第二個：進出貨每天都會用，跟盤點放在一起；手機底部 4 個還放得下（iOS 最多 5 個）
const TABS = [
  { id: 'home', label: '盤點', icon: 'clipboard' },
  { id: 'pick', label: '點貨', icon: 'list-check', fresh: true },
  { id: 'items', label: '品項', icon: 'box' },
  { id: 'lookup', label: '查型號', icon: 'search' },
  { id: 'settings', label: '設定', icon: 'settings', wide: true },
]
/**
 * 手機：底部分頁列（只在四個主頁）。平板、電腦（≥768px）：同一個元素變成右邊側邊欄，每一頁都有（sub＝子頁面，手機不顯示）。
 * fresh：新加的入口，標綠色「新」（手機上變成圖示右上角的小綠點）
 * 4.7.3 液態玻璃：手機是浮起來的膠囊、只放圖示（像 Instagram）；目前那一頁的圖示後面有一顆淡灰「鏡片」（.tab-lens），換頁時滑過去
 */
const tabBar = (active, sub = false) => {
  return `<nav class="tabbar${sub ? ' sub' : ''}" aria-label="主選單"><div class="side-head wide-only"><img class="brand-logo" src="logo.svg" alt=""><span class="side-name">聖佳智慧庫存</span><button class="side-toggle" data-action="side-toggle" aria-label="${sideCollapsed() ? '展開選單' : '收合選單'}" title="${sideCollapsed() ? '展開選單' : '收合選單'}">${icon('chevron-right', 20)}</button></div><div class="inner">${TABS.map((t) => `<button data-go="${t.id}" ${t.wide ? 'class="wide-only"' : ''} ${t.id === active ? 'aria-current="page"' : ''} aria-label="${t.label}${t.fresh ? '（新）' : ''}" title="${t.label}${t.fresh ? '（新）' : ''}">${icon(t.icon, 26)}<span>${t.label}</span>${t.fresh ? '<em class="tab-new" aria-hidden="true">新</em>' : ''}</button>`).join('')}</div></nav>`
}
// ───────── 液態玻璃鏡片：手機分頁列（橫）、平板／電腦側邊欄（直）、分段控制（.seg）共用這一份 ─────────
// 10/10 使用者：「液態玻璃那種如果滑過去，有沒有辦法做那種 iPhone 滑過去的效果」「各個載具都要有」
// - 點一下：鏡片用彈簧滑過去（有一點 overshoot）。按住拖：鏡片跟著手指／滑鼠走，放開彈到最近一格再換頁
// - 果凍拉伸：越快越長（順著方向最多 1.25 倍、另一邊最少 0.9 倍），停下來彈回 1；放大鏡：鏡片經過的圖示放大
// - 只改 transform（不重排版面、不重畫 App），放開才換頁；「減少動態效果」：沒有彈簧、拉伸、放大，直接切換
// - 做不到：iPhone 的網頁沒有震動回饋；Safari 做不到玻璃的折射扭曲
// 畫面每次重畫都是新的元素 → 狀態記在 LENSES（用 key 分），liquidLens() 每次重畫後接上新的元素
const reduceMotion = matchMedia('(prefers-reduced-motion: reduce)')
const phoneBar = matchMedia('(max-width: 767px)')
const LENSES = new Map()
let lensClickOk = false // 放開後程式自己按的那一下（換頁）不要擋
let lensEatClick = 0 // 拖完放開，瀏覽器補發的 click 不要算（不然會按到手指底下那一顆）
/** box：放鏡片的容器；items：可以選的按鈕；opts：axis 'x'|'y'、fit 'h'|'wh'（鏡片高／寬跟按鈕一樣）、mag 圖示放大倍數、isActive */
function liquidLens(key, box, items, opts) {
  let S = LENSES.get(key)
  if (!S) LENSES.set(key, (S = { idx: -1, pos: 0, v: 0, target: 0, st: 1, stv: 0, raf: 0, last: 0, drag: null, after: null, homes: [] }))
  let el = box.querySelector(':scope > .tab-lens')
  if (!el) {
    el = document.createElement('i')
    el.className = 'tab-lens'
    el.setAttribute('aria-hidden', 'true')
    box.prepend(el)
  }
  box.classList.add('has-lens')
  const axisChanged = S.axis !== opts.axis
  Object.assign(S, { box, el, items, axis: opts.axis, mag: opts.mag || 0 })
  const active = items.findIndex(opts.isActive)
  if (active < 0) {
    el.classList.add('off')
    S.idx = -1
    S.drag = null
    S.after = null
    return
  }
  el.classList.remove('off')
  const X = opts.axis === 'x'
  const a = items[active]
  if (opts.fit) {
    el.style.height = a.offsetHeight + 'px'
    if (opts.fit === 'wh') el.style.width = a.offsetWidth + 'px'
  }
  const len = X ? el.offsetWidth : el.offsetHeight
  S.len = len
  S.homes = items.map((b) => (X ? b.offsetLeft + (b.offsetWidth - len) / 2 : b.offsetTop + (b.offsetHeight - len) / 2))
  const prev = S.idx
  S.idx = active
  if (reduceMotion.matches || prev < 0 || axisChanged) {
    cancelAnimationFrame(S.raf)
    Object.assign(S, { raf: 0, v: 0, st: 1, stv: 0, after: null, drag: null, pos: S.homes[active], target: S.homes[active] })
  } else if (prev !== active) {
    S.target = S.homes[active] // 從現在的位置用彈簧滑過去
    lensRun(S)
  } else if (!S.raf && !S.drag) S.pos = S.target = S.homes[active]
  lensPaint(S)
}
function lensPaint(S) {
  if (!S.el) return
  const X = S.axis === 'x'
  const along = S.st
  const cross = Math.max(0.9, 1 - (S.st - 1) * 0.4)
  S.el.style.transform = `translate3d(${X ? S.pos.toFixed(2) : 0}px,${X ? 0 : S.pos.toFixed(2)}px,0)` + (Math.abs(along - 1) > 0.002 ? ` scale(${(X ? along : cross).toFixed(3)},${(X ? cross : along).toFixed(3)})` : '')
  if (!S.mag) return
  // 放大鏡：拖的時候一直有；點一下滑過去時越接近終點越小，停下來回到 1
  const pitch = S.homes.length > 1 ? Math.abs(S.homes[1] - S.homes[0]) || 1 : 1
  const act = S.drag?.moved ? 1 : Math.min(1, Math.abs(S.v) / 600 + Math.abs(S.pos - S.target) / pitch)
  S.items.forEach((b, k) => {
    const ic = b.querySelector('.ico')
    if (!ic) return
    const s = 1 + S.mag * Math.max(0, 1 - Math.abs(S.homes[k] - S.pos) / pitch) * act
    ic.style.transform = s > 1.002 ? `scale(${s.toFixed(3)})` : ''
  })
}
function lensStep(S, t) {
  const dt = S.last ? Math.min(0.034, (t - S.last) / 1000) : 1 / 60
  S.last = t
  // 簡單的彈簧：拖的時候硬一點（跟得上手指），放開後軟一點（有一點 overshoot）
  const [k, c] = S.drag?.moved ? [900, 52] : [320, 24]
  S.v += (-k * (S.pos - S.target) - c * S.v) * dt
  S.pos += S.v * dt
  const sp = Math.abs(S.v)
  S.stv += (-450 * (S.st - (1 + Math.min(0.25, sp / 2400))) - 16 * S.stv) * dt
  S.st = Math.min(1.25, Math.max(0.94, S.st + S.stv * dt))
  const near = Math.abs(S.pos - S.target) < 0.8 && sp < 15
  if (near && S.after) {
    const f = S.after
    S.after = null
    f()
  }
  if (!S.drag && near && Math.abs(S.st - 1) < 0.003 && Math.abs(S.stv) < 0.05) {
    Object.assign(S, { pos: S.target, v: 0, st: 1, stv: 0, raf: 0, last: 0 })
    lensPaint(S)
    return
  }
  lensPaint(S)
  S.raf = requestAnimationFrame((tt) => lensStep(S, tt))
}
function lensRun(S) {
  if (S.raf) return
  S.last = 0
  S.raf = requestAnimationFrame((t) => lensStep(S, t))
}
/** 每次畫完畫面：分頁列／側邊欄、每一個分段控制接上鏡片 */
function mountLenses(root = $app) {
  const nav = root.querySelector('.tabbar .inner')
  if (nav) {
    const phone = phoneBar.matches
    const shown = nav.offsetParent !== null // 手機的子頁面沒有分頁列
    const items = shown ? [...nav.querySelectorAll(':scope > button')].filter((b) => !phone || !b.classList.contains('wide-only')) : []
    liquidLens('nav', nav, items, { axis: phone ? 'x' : 'y', fit: phone ? null : 'h', mag: phone ? 0.18 : 0.12, isActive: (b) => b.hasAttribute('aria-current') })
  }
  root.querySelectorAll('.seg').forEach((seg, i) => {
    liquidLens(`seg:${state.view}:${i}`, seg, [...seg.querySelectorAll(':scope > button')], { axis: 'x', fit: 'wh', isActive: (b) => b.getAttribute('aria-selected') === 'true' })
  })
}
document.addEventListener('pointerdown', (e) => {
  if (reduceMotion.matches || e.button > 0 || !e.isPrimary) return
  const box = e.target.closest?.('.has-lens')
  const S = box && [...LENSES.values()].find((s) => s.box === box && s.idx >= 0)
  if (!S) return
  S.drag = { id: e.pointerId, x0: e.clientX, y0: e.clientY, r: box.getBoundingClientRect(), moved: false }
})
addEventListener(
  'pointermove',
  (e) => {
    for (const S of LENSES.values()) {
      const d = S.drag
      if (!d || d.id !== e.pointerId) continue
      const X = S.axis === 'x'
      const off = X ? e.clientX - d.x0 : e.clientY - d.y0
      if (!d.moved) {
        if (Math.abs(off) < 8) continue
        // 分段控制在內容裡：手指主要是上下滑 → 交給頁面捲動
        if (Math.abs(X ? e.clientY - d.y0 : e.clientX - d.x0) > Math.abs(off)) {
          S.drag = null
          continue
        }
        d.moved = true
        S.box.classList.add('lens-drag')
      }
      let p = (X ? e.clientX - d.r.left : e.clientY - d.r.top) - S.len / 2
      const lo = S.homes[0]
      const hi = S.homes[S.homes.length - 1]
      if (p < lo) p = lo - (lo - p) * 0.3
      else if (p > hi) p = hi + (p - hi) * 0.3
      S.target = p
      lensRun(S)
    }
  },
  { passive: true },
)
function lensEnd(e) {
  for (const S of LENSES.values()) {
    const d = S.drag
    if (!d || d.id !== e.pointerId) continue
    S.drag = null
    S.box?.classList.remove('lens-drag')
    if (!d.moved) continue
    lensEatClick = performance.now() + 450
    let k = S.idx
    if (e.type === 'pointerup') {
      let best = Infinity
      S.homes.forEach((h, i) => {
        if (Math.abs(h - S.target) < best) (best = Math.abs(h - S.target)), (k = i)
      })
    }
    S.target = S.homes[k]
    if (k !== S.idx)
      S.after = () => {
        lensClickOk = true
        try {
          S.items[k]?.click()
        } finally {
          lensClickOk = false
        }
      }
    lensRun(S)
  }
}
addEventListener('pointerup', lensEnd)
addEventListener('pointercancel', lensEnd)
addEventListener(
  'click',
  (e) => {
    if (lensClickOk || performance.now() > lensEatClick || !e.target.closest?.('.has-lens')) return
    lensEatClick = 0
    e.preventDefault()
    e.stopImmediatePropagation()
  },
  true,
)
let lensResize = 0
addEventListener('resize', () => {
  cancelAnimationFrame(lensResize)
  lensResize = requestAnimationFrame(() => mountLenses())
})
// 上方標題列：捲下去才出現玻璃的分隔線（在最上面時跟背景融在一起，跟 iOS 一樣）
let scrolledFlag = false
addEventListener(
  'scroll',
  () => {
    const s = scrollY > 4
    if (s !== scrolledFlag) document.documentElement.classList.toggle('is-scrolled', (scrolledFlag = s))
  },
  { passive: true },
)
/** 電腦、平板的右側選單收合（只剩圖示）；記在這台 */
// 沒按過收合鈕：平板直放（不到 1000px 寬）預設收起來，內容才不會被擠到只剩七成
const sideCollapsed = () => {
  const v = ls.get(LS.sideCollapsed)
  return v ? v === '1' : innerWidth < 1000
}
document.body.classList.toggle('side-collapsed', sideCollapsed())
addEventListener('resize', () => {
  const was = document.body.classList.contains('side-collapsed')
  if (was === sideCollapsed()) return
  document.body.classList.toggle('side-collapsed', !was)
  const b = document.querySelector('.side-toggle')
  if (b) b.setAttribute('aria-label', !was ? '展開選單' : '收合選單')
})
/** 子頁面屬於哪一個主頁（側邊欄標哪一個） */
const TAB_OF = { capture: 'home', analyzing: 'home', review: 'home', report: 'home', quality: 'home', item: 'items', locations: 'items', catalog: 'items', lookup: 'lookup', settings: 'settings', pick: 'pick', 'pick-edit': 'pick', 'pick-result': 'pick' }

/**
 * 首頁的大數字方塊（4.6）：今天盤了幾件、要確認幾個、該叫貨幾種、要複盤幾項；點了直接去處理。
 * 要確認＝今天的、或還沒按完成的盤點裡，還沒看過的「要確認」框。
 */
function homeStats(sessions, items) {
  const today = sessions.filter((s) => inRange(s, 'today'))
  const recent = sessions.filter((s) => !s.linkedAt || inRange(s, 'today'))
  const doubtful = recent.map((s) => [s, doubtsOf(s).length]).filter(([, n]) => n)
  const doubts = doubtful.reduce((n, [, k]) => n + k, 0)
  const order = items.filter(needsOrder).length
  const recount = items.filter(needsRecount).length
  const tile = ({ n, unit, label, tone, attrs }) => {
    const tag = attrs && n ? 'button' : 'div'
    return `<${tag} class="tile ${n ? tone : 'zero'}" ${n ? attrs || '' : ''}><span class="tile-num">${n}<small>${unit}</small></span><span class="tile-label">${label}</span>${tag === 'button' ? chev : ''}</${tag}>`
  }
  return `<p class="section-title">現在要處理的 <span class="badge new">新</span></p><div class="home-stats" role="group" aria-label="現在要處理的">
    ${tile({ n: reportOf(today).total, unit: '件', label: `今天盤了・${today.length} 次`, tone: 'brand', attrs: 'data-go="report"' })}
    ${tile({ n: doubts, unit: '個', label: '要確認', tone: 'warn', attrs: doubtful.length && canEdit() ? `data-open-doubts="${esc(doubtful[0][0].id)}"` : '' })}
    ${tile({ n: order, unit: '種', label: '該叫貨', tone: 'bad', attrs: 'data-go="items" data-item-filter="order"' })}
    ${tile({ n: recount, unit: '項', label: '要複盤', tone: 'warn', attrs: 'data-recount-list' })}
  </div>`
}

async function viewHome() {
  const sessions = await db.all()
  const hasKey = hasAi()
  const items = await itemsAll()
  // 多選：按「選取」後，每一筆前面出現圓圈，下面出現動作列（匯出、記進品項庫、刪除）
  const sel = !!state.selecting
  state.selected ??= new Set()
  for (const id of [...state.selected]) if (!sessions.some((s) => s.id === id)) state.selected.delete(id)
  const picked = sessions.filter((s) => state.selected.has(s.id))
  const unfinished = sessions.filter((s) => !s.linkedAt).length
  return `
  <main class="app">
    <div class="nav">${syncReady() ? `<button class="btn small plain sync-pill ${state.syncState || ''}" data-action="sync-now">${icon('cloud', 18)}<span class="sync-text">${esc(syncLabel())}</span></button>` : '<span></span>'}<button class="icon-btn" data-go="settings" aria-label="設定">${icon('settings', 24)}</button></div>
    <h1 class="large-title brand"><img class="brand-logo" src="logo.svg" alt="聖佳 LOGO">聖佳智慧庫存</h1>
    <div class="home-grid"><div class="home-main">
    <div class="home-cards">
    ${
      !canEdit()
        ? `<div class="hint-card"><b>你是檢視者（只能看）</b>：可以看大家的盤點紀錄、品項庫，也可以下載 Excel；不能盤點或修改。需要盤點請找管理員改成「編輯者」。</div>`
        : hasKey
        ? `<button class="hero-btn" data-action="new"><span class="hero-icon" aria-hidden="true">${icon('camera', 30)}</span><span class="grow"><b>新盤點</b><span class="meta">拍一格貨架，AI 幫你數<br>拍到 ${locTag(locations()[0]?.code || 'A-03', 'sm')} 會自動填位置</span></span>${chev}</button>`
        : syncReady()
          ? `<div class="setup-card"><span class="setup-ico" aria-hidden="true">${icon('camera', 26)}</span><div class="grow"><b>還不能拍照盤點：擁有者還沒把 AI 金鑰放到雲端</b><p>請擁有者到「設定 → Gemini API Key」按「放到雲端」，大家就能直接用，不用自己申請。<br>急著用：也可以到「設定」貼上自己的 Gemini API Key。</p><button class="btn small" data-go="settings">去設定</button></div></div>`
          : `<div class="setup-card"><span class="setup-ico" aria-hidden="true">${icon('camera', 26)}</span><div class="grow"><b>先設定 AI，才能拍照盤點</b><p>到「設定」貼上 Gemini API Key<span class="nb">（只存在這台）</span>；或請擁有者把金鑰放到雲端共用。<br>只想看別台盤點的結果：到「設定 → 我收到連結碼了」貼上邀請連結就好，不用 Key。</p><button class="btn small" data-go="settings">去設定</button></div></div>`
    }
    </div>
    ${sessions.length || items.length ? homeStats(sessions, items) : ''}
    </div><div class="home-side">
    ${
      sessions.length
        ? `<div class="group home-report"><button class="row" data-go="report"><span class="row-ico" aria-hidden="true">${icon('chart')}</span><span class="grow"><span class="title">總表與匯出</span><br><span class="meta">很多次盤點合起來看；<span class="nb">下載 Excel</span>${ls.get(LS.sheet) ? '、<span class="nb">同步試算表</span>' : ''}</span></span>${chev}</button></div>`
        : ''
    }
    <div class="list-head"><p class="section-title">盤點紀錄${sel ? `・已選 ${picked.length} 筆` : ''}</p>${sessions.length ? `<button class="btn small plain" data-action="${sel ? 'select-done' : 'select-start'}">${sel ? '完成' : '選取'}</button>` : ''}</div>
    ${
      sel
        ? `<div class="chips select-chips"><button class="chip" data-action="select-all">${picked.length === sessions.length ? '全不選' : '全選'}</button>${unfinished ? `<button class="chip" data-action="select-unfinished">選還沒按完成的（${unfinished}）</button>` : ''}</div>`
        : ''
    }
    ${
      sessions.length
        ? `<div class="group">${sessions
            .map((s) => {
              const failed = s.photos.length && s.photos.every((p) => p.status !== 'done')
              const meta = [byName(s) ? `${byName(s)} 盤・${fmtTime(s.createdAt)}` : fmtTime(s.createdAt), s.photos.length > 1 ? `${s.photos.length} 張照片` : '', s.syncedAt ? '已同步到試算表' : ''].filter(Boolean).join('・')
              const on = sel && state.selected.has(s.id)
              const sessImg = s.photos[0]?.blob ? `<img class="sess-thumb" src="${urlOf(s.photos[0])}" alt="" loading="lazy" decoding="async">` : ''
              // 縮圖點一下：放大看這次的照片（好幾張可以左右換）；多選的時候點哪裡都是勾選
              const sessZoom = () => zoomWrap(sessImg, zoomAttrs(s.photos.filter((p) => p.blob).map(urlOf).join('\n'), [summaryOf(s), fmtTime(s.createdAt)].filter(Boolean).join('・'), { places: s.place ? [s.place] : [], inRow: true }), 'md')
              const body = `${sessImg ? (sel ? sessImg : sessZoom()) : `<span class="thumb-lost ${s.photos[0]?.pending ? 'pending' : ''}" aria-hidden="true"></span>`}<span class="grow">${s.place ? `<span class="sess-place">${placeHtml(s.place, { name: false })}</span>` : ''}<span class="title">${esc(failed ? '沒有辨識成功（點進去再試一次）' : summaryOf(s))}</span><br><span class="meta">${esc(meta)}${s.linkedAt ? '' : `${meta ? '・' : ''}<span class="unfinished">還沒按完成</span>`}</span></span>`
              return sel
                ? `<button class="row ${on ? 'picked' : ''}" data-pick="${esc(s.id)}" aria-pressed="${on}"><span class="pick" aria-hidden="true">${on ? icon('check', 16) : ''}</span>${body}</button>`
                : `<button class="row" data-open="${esc(s.id)}">${body}${chev}</button>`
            })
            .join('')}</div>`
        : `<div class="empty"><span class="empty-ico" aria-hidden="true">${icon('box', 36)}</span><p><b>還沒有盤點紀錄</b><br>${!canEdit() ? '別台盤點完、同步後，紀錄會出現在這裡。' : hasKey ? '按「新盤點」，拍一層貨架試試看。' : '按「去設定」設定好 AI，就能拍照盤點；別台同步過來的紀錄也會出現在這裡。'}</p></div>`
    }
    </div></div>
  </main>
  ${
    sel
      ? `<div class="toolbar select-bar"><div class="inner">
          <button class="btn secondary" data-action="bulk-export" ${picked.length ? '' : 'disabled'}>匯出 Excel</button>
          <button class="btn secondary edit-only" data-action="bulk-finish" ${picked.length ? '' : 'disabled'}>記進品項庫</button>
          <button class="btn danger edit-only" data-action="bulk-delete" ${picked.length ? '' : 'disabled'}>刪除</button>
        </div></div>`
      : tabBar('home')
  }`
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
    <div class="group"><label class="row"><span style="width:96px" class="muted">位置</span><input class="inline place-input" id="place" placeholder="${locations().length ? '選下面的儲位或打代號' : '例：A-01（會自動填）'}" value="${esc(s.place)}" autocomplete="off"></label></div>
    ${
      locations().length
        ? `<div class="chips loc-chips" role="group" aria-label="選儲位">${locations()
            .map((l) => `<button class="chip" data-place="${esc(l.code)}" aria-pressed="${canon(l.code) === canon(s.place)}">${locTag(l.code, 'sm')}${l.name ? `<small>${esc(l.name)}</small>` : ''}</button>`)
            .join('')}</div>`
        : `<p class="footnote">到「品項 → 儲位」建立代號、印標籤貼在貨架上：之後拍照會自動填位置，重盤同一格也會自動更新數量。</p>`
    }
    <p class="section-title">照片（${s.photos.length}）</p>
    ${
      s.photos.length
        ? `<div class="photo-strip cap-strip">${s.photos.map((p, i) => `<div class="cap-thumb"><button class="cap-view" data-preview-photo="${i}" aria-label="放大看第 ${i + 1} 張"><img src="${urlOf(p)}" alt="" decoding="async"><span class="zoom-badge" aria-hidden="true">${icon('zoom-in', 13)}</span></button><button class="cap-del" data-remove-photo="${i}" aria-label="刪掉第 ${i + 1} 張" title="刪掉這張"><span>${icon('x', 14)}</span></button></div>`).join('')}</div><p class="footnote">點縮圖放大看；拍錯了按 × 刪掉，可以復原。</p>`
        : ''
    }
    <div class="row-actions" style="margin-top:12px">
      <label class="btn secondary" style="flex:1">${icon('camera')}拍照<input type="file" accept="image/*" capture="environment" id="cam" class="sr-only"></label>
      <label class="btn secondary" style="flex:1">${icon('image')}從相簿選<input type="file" accept="image/*" multiple id="pick" class="sr-only"></label>
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
  const errors = s.photos.map((p, i) => (p.status !== 'done' ? `<div class="error-card">第 ${i + 1} 張沒辨識成功：${esc(p.error || '還沒辨識（被取消）')} <button class="btn small secondary" data-retry="${i}" style="margin-left:6px">再試一次</button>${p.errorDetail ? `<div style="margin-top:8px;font-size:13px;color:var(--text-2);word-break:break-all">錯誤代碼：${esc(p.errorDetail)}</div>` : ''}</div>` : '')).join('')
  const notes = s.photos.map((p, i) => (p.note ? `<p class="footnote">第 ${i + 1} 張 AI 備註：${esc(p.note)}</p>` : '')).join('')
  return `
  <main class="app">
    <div class="nav">${backBtn('home', '盤點')}<button class="btn small secondary edit-only" data-action="add">${icon('plus', 18)}手動新增</button></div>
    <h1 class="large-title${s.place ? ' place-title' : ''}">${s.place ? placeHtml(s.place, { size: 'xl' }) : '盤點結果'}</h1>
    ${
      byName(s) || syncReady()
        ? `<div class="by-line"><span class="avatar small" aria-hidden="true">${esc((byName(s) || '?').slice(0, 1))}</span><span class="grow"><b>盤點人：${esc(byName(s) || '沒有記錄')}</b><br><span class="meta">${fmtTime(s.createdAt)}${s.model ? `・${esc(s.model)}` : ''}</span></span>${canManage() && syncReady() ? '<button class="btn small plain" data-action="counter-session">更正</button>' : ''}</div>`
        : `<p class="subtitle">${fmtTime(s.createdAt)}${s.model ? `・${esc(s.model)}` : ''}</p>`
    }
    ${errors ? `<div class="stack">${errors}</div>` : ''}
    <div class="review-cols"><div class="review-top">
    ${
      groups.length
        ? `<section class="summary" aria-label="盤點總結">
            <div class="sum-nums">
              <div><span class="sum-big">${total}</span><span class="sum-unit">件</span></div>
              <div class="sum-side"><b>${groups.length}</b> 種${s.photos.length > 1 ? `・${s.photos.length} 張照片` : ''}</div>
            </div>
            ${
              doubts.length
                ? `<button class="sum-doubt edit-only" data-action="review-doubts"><span class="sum-dot" aria-hidden="true">?</span><span class="grow"><b>${doubts.length} 個要確認</b><br><span class="meta">可能尺寸不同，或 AI 沒把握</span></span>${chev}</button>`
                : `<div class="sum-ok">${icon('check', 20)}沒有需要確認的${state.refining && state.refining.session === s.id ? '（相似品還在比對）' : ''}</div>`
            }
            ${
              s.eval?.types.length
                ? `<button class="sum-eval" data-go="quality"><span class="grow">AI 這次：${s.eval.types.length} 種裡 <b>${s.eval.types.filter((t) => t.ai === t.final).length}</b> 種數量一個不差${s.eval.missed || s.eval.extra || s.eval.misclass ? `<br><span class="meta">${[s.eval.missed && `漏數 ${s.eval.missed}`, s.eval.extra && `多數 ${s.eval.extra}`, s.eval.misclass && `分錯 ${s.eval.misclass}`].filter(Boolean).join('・')}</span>` : ''}</span>${chev}</button>`
                : ''
            }
          </section>`
        : ''
    }
    </div><div class="review-photo">
    ${
      photo
        ? `${
            s.photos.some((p) => p.lost && !fetching.has(s.id))
              ? `<div class="hint-card lost-card" role="status"><b>${s.photos.filter((p) => p.lost).length} 張照片在這台手機不見了</b>（iPhone 的問題；框和數量都還在）。雲端有備份的話可以拿回來。<button class="btn small" data-action="repair-photos">從雲端拿回照片</button></div>`
              : ''
          }<div class="photo-wrap ${state.focus || state.focusObj ? 'focus' : ''} ${state.addMode ? 'adding' : ''}" data-photo style="--ar:${photo.w && photo.h ? (photo.w / photo.h).toFixed(3) : '1.333'}">${
            photo.blob ? `<img src="${urlOf(photo)}" alt="第 ${state.photoIndex + 1} 張照片">` : `<div class="photo-ph ${fetching.has(s.id) || photo.pending ? 'loading' : ''}" role="img" aria-label="${photo.pending ? '照片下載中' : '這張照片不見了'}">${fetching.has(s.id) || photo.pending ? '<span class="photo-ph-text">照片下載中…</span>' : ''}</div>`
          }${boxes}
             <button class="photo-zoom" data-action="zoom" aria-label="放大看照片">${icon('maximize', 22)}</button>
           </div>
           ${state.addMode ? '<div class="add-hint" role="status"><b>點照片上漏掉的那一個</b>，會在那裡加一個框 <button class="btn small plain" data-action="add-cancel">取消</button></div>' : ''}
           ${s.photos.length > 1 ? `<div class="photo-strip">${s.photos.map((p, i) => `<button class="${i === state.photoIndex ? 'on' : ''} ${p.blob ? '' : 'lost'}" data-photo-index="${i}" aria-label="看第 ${i + 1} 張${p.pending ? '（下載中）' : p.lost ? '（照片不見了）' : ''}">${p.blob ? `<img src="${urlOf(p)}" alt="" decoding="async">` : ''}</button>`).join('')}</div>` : ''}
           <div class="row-actions edit-only" style="margin-top:10px"><button class="btn small secondary" data-action="add-box" ${state.addMode ? 'disabled' : ''}>${icon('plus', 18)}漏掉的，點照片補一個</button></div>
           <p class="footnote">點照片上的框：選它是哪一種，改好會關掉視窗，下面跳出「復原」可以改回來。點品項清單：看那一種在哪裡；數量不對按 ＋／－。</p>`
        : ''
    }
    </div><div class="review-side">
    ${
      state.viewer && photo
        ? `<div class="viewer" role="dialog" aria-label="放大看照片">
            <div class="viewer-bar">
              <span class="zoom-label">${reviewZoomLabel()}</span>
              <button class="btn small secondary" data-action="zoom-out" ${state.zoom <= 1 ? 'disabled' : ''} aria-label="縮小">${icon('minus', 20)}</button>
              <button class="btn small secondary" data-action="zoom-in" ${state.zoom >= ZOOM_MAX ? 'disabled' : ''} aria-label="放大">${icon('plus', 20)}</button>
              <button class="viewer-close" data-action="zoom-close" aria-label="關閉">${icon('x', 22)}</button>
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
              <span class="name">${esc(g.label)}</span><span class="pencil" aria-hidden="true">${icon('edit', 16)}</span>${g.manual ? '<span class="badge edit">手動</span>' : doubtsByKey.get(g.key) ? `<span class="badge low">${doubtsByKey.get(g.key)} 個要確認</span>` : ''}${g.edited && !g.manual ? '<span class="badge edit">已修正</span>' : ''}${itemsCache && !findItem(itemsCache, g) ? '<span class="badge ok">新品項</span>' : ''}
              <br><span class="spec">${detailOf(g) ? detailHtml(g) : '點名稱補品牌、型號、尺寸'}${!g.manual && g.count !== g.boxes ? `・<span class="nb">照片裡 ${g.boxes} 個</span>` : ''}</span>
            </button>
            <span class="stepper edit-only"><button data-step="-1" data-key="${esc(g.key)}" aria-label="減一">${icon('minus', 20)}</button><input inputmode="numeric" value="${g.count}" data-count="${esc(g.key)}" aria-label="${esc(g.label)} 數量"><button data-step="1" data-key="${esc(g.key)}" aria-label="加一">${icon('plus', 20)}</button></span><span class="qty view-only"><b>${g.count}</b></span>
          </div>`,
            )
            .join('')}</div>`
        : `<div class="empty"><p>這次沒有找到商品。<br>可以重拍，或按右上「手動新增」。</p></div>`
    }
    ${
      groups.length
        ? `<label class="row golden-row edit-only"><span class="grow"><span class="title">當成標準答案</span><br><span class="meta">兩個人各自數過、數字一樣才勾。之後在「AI 準不準」可以拿來考 AI（換模型、改設定時看有沒有變差）。</span></span><input type="checkbox" id="golden" ${s.golden ? 'checked' : ''} style="width:22px;height:22px"></label>`
        : ''
    }
    <div class="row-actions edit-only" style="margin-top:22px">
      <button class="btn danger small" data-action="delete-session">刪除這次盤點</button>
    </div>
    </div></div>
  </main>
  <div class="toolbar"><div class="inner"><button class="btn secondary" data-action="export">匯出</button><button class="btn edit-only" data-action="finish">完成・記進品項庫</button></div></div>`
}

/** 設定頁：共用金鑰（擁有者放到雲端／同事看到「正在用共用的」）；都不適用就什麼都不顯示 */
function aiShareRows(key) {
  if (!syncReady()) return ''
  const rows = []
  const fresh = '<span class="badge new">新</span>'
  // 擁有者只看下面「金鑰已放在雲端給大家共用」那一列就好（不然兩列說的是同一件事）
  if (aiSharedOn() && !key && myRole() !== 'owner')
    rows.push(
      canEdit()
        ? `<div class="row"><span class="grow"><span class="title ok-line">${icon('check', 18)}正在用公司共用的金鑰</span>${fresh}<br><span class="meta">擁有者把金鑰放在雲端（Apps Script），你的手機上沒有金鑰，可以直接拍照辨識；每次會多 1～2 秒。</span></span></div>`
        : `<div class="row"><span class="grow"><span class="title">你是檢視者，不能拍照辨識</span>${fresh}<br><span class="meta">檢視者只能看盤點紀錄、品項庫和下載 Excel。需要盤點，請找擁有者或管理員改成「編輯者」。</span></span></div>`,
    )
  if (myRole() === 'owner')
    rows.push(
      ls.get(LS.aiShared) === '1'
        ? `<div class="row"><span class="grow"><span class="title">金鑰已放在雲端給大家共用</span>${fresh}<br><span class="meta">同事不用跟你要金鑰；大家的用量都算在這把金鑰上。</span></span><button class="btn small secondary" data-action="ai-unshare">停止共用</button></div>`
        : `<div class="row"><span class="grow"><span class="title">把金鑰放到雲端，大家共用</span>${fresh}<br><span class="meta">金鑰存在你的 Apps Script，同事的手機不會有，不用一個一個給。</span></span><button class="btn small" data-action="ai-share">放到雲端</button></div>`,
    )
  return rows.length ? `<div class="group">${rows.join('')}</div>` : ''
}

/**
 * 設定頁的「連線測試」：擁有者放左欄；其他人右欄（沒有「Google 試算表」那一區，右欄比較短，搬過去兩欄才一樣高）。
 * 沒有 AI 可以用（例如檢視者）就不顯示：按了也只會失敗。
 */
const diagSection = () =>
  hasAi()
    ? `<p class="section-title">連線測試</p>
    <div class="stack"><button class="btn small secondary" data-action="diagnose">測試連線</button><div id="diag"></div></div>
    <p class="footnote">辨識一直失敗時按這個，把結果截圖傳給擁有者（管理這個 App 的人）。</p>`
    : ''
/** 還沒設定任何同步、也沒有金鑰的新手機（大多是剛收到邀請的同事） */
const freshPhone = () => !syncReady() && !ls.get(LS.sheet) && !ls.get(LS.key)
async function viewSettings() {
  const key = ls.get(LS.key)
  const model = ls.get(LS.model)
  const samples = await idb.samples.all().catch(() => [])
  return `
  <main class="app">
    <div class="nav">${backBtn('home', '盤點')}</div>
    <h1 class="large-title">設定</h1>
    ${
      // 還沒設定任何同步的新手機：大多是被邀請的同事 →「我收到連結碼了」放最前面（以前要往下捲很久）
      freshPhone()
        ? `<div class="hint-card stack join-first"><div><b>同事傳了邀請連結給你？</b> <span class="badge new">新</span><br>在 LINE 直接<b>點邀請連結</b>就會自動加入；點了沒反應，把整段連結貼在這裡：</div>
        <input class="field" id="link-code" placeholder="貼上收到的邀請連結" autocomplete="off" spellcheck="false">
        <button class="btn small" data-action="sync-link">加入</button>
        <p class="footnote" style="margin:0">自己就是擁有者（管理這個 App 的人）：不用填這裡，往下設定 Gemini API Key、Google 試算表。</p></div>`
        : ''
    }
    <div class="settings-cols"><div class="settings-col">
    <p class="section-title">Gemini API Key</p>
    ${(() => {
      const keyForm = `<input class="field" id="apikey" type="password" placeholder="貼上 API Key（AIza 開頭）" value="${esc(key)}" autocomplete="off" spellcheck="false">
      <div class="row-actions"><button class="btn small" data-action="save-key">儲存並測試</button><button class="btn small secondary" data-action="toggle-key">顯示／隱藏</button></div>
      <p class="footnote" style="margin:0">沒有 Key？到 <a href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer" style="color:var(--tint)">Google AI Studio</a> 免費建立。Key 只存在這支手機，照片只會送到 Google Gemini 分析。</p>`
      // 已經在用擁有者的共用金鑰：自己貼金鑰收進「進階」，同事才不會以為要自己去申請
      return aiSharedOn() && !key
        ? `<div class="stack">${aiShareRows(key)}${canEdit() ? `<details class="steps"><summary>進階：改用自己的金鑰</summary><div class="stack" style="margin-top:6px">${keyForm}</div></details>` : ''}</div>`
        : `<div class="stack">${keyForm}${aiShareRows(key)}</div>`
    })()}
    <p class="section-title">辨識模型</p>
    <div class="group"><div class="row"><span class="grow"><span class="title">${esc(model || '自動挑選')}</span><br><span class="meta">自動挑能看圖、最新又快的 Flash；被下架會自動換。相似品分不開時，可以改用 Pro</span></span>${hasAi() ? '<button class="btn small secondary" data-action="pick-model">重新挑選</button>' : ''}</div>
      <button class="row" data-go="quality"><span class="grow"><span class="title" style="color:var(--tint)">AI 準不準</span><br><span class="meta">AI 原本數的跟你確認後的比；也可以用標準答案考 AI</span></span>${chev}</button></div>
    <div id="models"></div>
    <p class="section-title">店內品項清單</p>
    <textarea class="field" id="catalog" spellcheck="false" ${canEdit() ? '' : 'readonly aria-readonly="true"'}>${esc(catalogLines().join('\n'))}</textarea>
    <p class="footnote">一行一種品項，括號裡寫規格或別名。AI 會照這裡的名稱寫，修正時也會跳出來給你選。${canEdit() ? '' : '（檢視者只能看）'}</p>
    <div class="row-actions edit-only" style="margin-top:10px"><button class="btn small" data-action="save-catalog">儲存清單</button><button class="btn small secondary" data-action="reset-catalog">恢復預設</button></div>
    <p class="section-title">樣品照（${samples.length}）</p>
    ${
      samples.length
        ? `<div class="group">${samples
            .map(
              (s, i) =>
                `<div class="row">${zoomWrap(`<img class="sample-thumb" src="${sampleUrl(s)}" alt="" loading="lazy" decoding="async">`, zoomAttrs(sampleUrl(s), ['樣品照', s.label, detailOf(s)].filter(Boolean).join('・'), { group: 'samples' }), 'md')}<span class="grow"><span class="title">${esc(s.label)}</span><br><span class="meta">${esc(detailOf(s) || '沒有填品牌、型號、尺寸')}${i >= MAX_SAMPLES ? '・超過 30 張，這張不會送' : ''}</span></span><button class="btn small danger edit-only" data-action="del-sample" data-id="${esc(s.id)}">刪除</button></div>`,
            )
            .join('')}</div>`
        : ''
    }
    <div class="row-actions edit-only" style="margin-top:10px"><label class="btn small secondary">${icon('camera', 20)}拍一張樣品<input type="file" accept="image/*" capture="environment" id="sample-cam" class="sr-only"></label></div>
    <p class="footnote">長得很像、只差尺寸的商品（例如不同分數的三通），每一種存一張樣品照，名稱和尺寸寫清楚。辨識時會一起送給 AI 比對（最多 30 張，新的優先）。<br>最快的存法：盤點結果裡先點那個框、再點一次 →「儲存，並存成樣品照」。</p>
    ${myRole() === 'owner' ? diagSection() : ''}
    </div><div class="settings-col">
    ${
      myRole() === 'owner'
        ? `<p class="section-title">Google 試算表</p>
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
      <p><b>以前連結過的：</b>要有「品項庫」工作表和多台同步，請重新複製程式碼貼上 → 存檔 →「部署」→「管理部署作業」→ 鉛筆（編輯）→ 版本選「新版本」→ 部署（網址不變；會再問一次授權，因為要存到你的雲端硬碟）。</p>
    </details>`
        : ''
    }
    <p class="section-title">共用（跟 Google 雲端硬碟一樣）</p>
    ${
      syncReady()
        ? `${
            serverOld()
              ? `<div class="hint-card stack" style="margin-bottom:12px"><div><b>Google 那邊的程式碼要更新</b>（新版同步只拿資料、照片打開才抓，快很多）。請<b>擁有者</b>用電腦做一次，約 3 分鐘：<ol class="steps-list" style="margin:8px 0 0"><li>按「複製試算表程式碼」。</li><li>打開當初那個 Google 試算表 → 擴充功能 → Apps Script → 全選、貼上取代 → 存檔。</li><li>右上「部署」→「<b>管理部署作業</b>」→ 鉛筆（編輯）→ 版本選「<b>新版本</b>」→ 部署。<br>（不要用「新增部署作業」，網址會變）</li></ol></div><button class="btn small" data-action="sheet-copy">複製試算表程式碼</button></div>`
              : ''
          }<div class="group">
            <div class="row"><span class="avatar" aria-hidden="true">${esc((ls.get(LS.memberName) || '我').slice(0, 1))}</span><span class="grow"><span class="title">${esc(ls.get(LS.memberName) || '我')}（這台）</span><br><span class="meta">${ROLE_LABEL[myRole()]}・<span class="sync-text">${state.syncState === 'error' && state.syncError ? esc(state.syncError) : esc(syncLabel())}</span></span></span><button class="btn small" data-action="sync-now">立即同步</button></div>
            <div class="row"><span class="grow"><span class="title">這台的盤點人</span><br><span class="meta">新盤點會自動記成這個人，不用每次選${canManage() ? '' : '；由擁有者或管理員設定'}</span></span>${
              canManage() ? `<button class="btn small secondary" data-action="counter-device">${esc(currentCounter()?.name || '選擇')}</button>` : `<b>${esc(currentCounter()?.name || '—')}</b>`
            }</div>
            ${
              canManage()
                ? `<button class="row" data-action="share-open"><span class="grow"><span class="title" style="color:var(--tint)">共用設定</span><br><span class="meta">邀請同事、改權限（只能看／可以改）、移除離職的人</span></span>${chev}</button>`
                : `<div class="row muted">${ROLE_DESC[myRole()]}。要加人或改權限，請找擁有者或管理員。</div>`
            }
            <button class="row" data-action="full-pull" ${repairing ? 'disabled' : ''}><span class="grow"><span class="title" style="color:var(--tint)">${repairing ? '正在從雲端重新下載…' : '從雲端重新下載全部'}</span><br><span class="meta">這台跟別台不一樣、少了紀錄或照片時用；這台比較新的不會被蓋掉</span></span>${chev}</button>
            <button class="row" data-action="sync-leave"><span class="grow"><span class="title" style="color:var(--red)">這台退出並清除資料</span><br><span class="meta">交還手機、換手機時用；雲端的資料不會刪</span></span>${chev}</button>
          </div>`
        : ls.get(LS.sheet)
          ? '<button class="btn block" data-action="sync-start">開啟多人同步（我是擁有者）</button><p class="footnote">開啟後到「共用設定」邀請同事：每個人一組自己的連結碼，可以設「只能看」或「可以改」。</p>'
          : '<div class="group"><div class="row muted">擁有者：先完成上面的 Google 試算表連結，再回來開啟同步。<br>被邀請的人：直接在下面貼上收到的連結碼。</div></div>'
    }
    ${
      syncReady() || freshPhone()
        ? ''
        : `<details class="steps"><summary>我收到連結碼了（被邀請的人用）</summary>
      <div class="stack" style="margin-top:8px">
        <p class="footnote" style="margin:0">最簡單：在 LINE 直接<b>點邀請連結</b>就會自動加入。點了沒反應，再把整段連結貼在這裡。</p>
        <input class="field" id="link-code" placeholder="貼上收到的邀請連結" autocomplete="off" spellcheck="false">
        <button class="btn small" data-action="sync-link">加入</button>
      </div>
    </details>`
    }
    <p class="section-title">盤點規則</p>
    <div class="group">
      <label class="row"><span class="grow"><span class="title">盲盤</span><br><span class="meta">盤點的人看不到帳面數、差異（擁有者、管理員看得到）。看得到帳面數，盤點就會變成「對答案」，抓不到錯。</span></span><input type="checkbox" id="rule-blind" ${ls.get(LS.blind, '1') === '1' ? 'checked' : ''} ${canManage() ? '' : 'disabled'} style="width:22px;height:22px"></label>
      <label class="row"><span class="grow"><span class="title">有差異先複盤</span><br><span class="meta">盤到的跟帳面不一樣：請另一個人再數一次，兩次一樣才確定，再選原因，由擁有者或管理員決定要不要調整帳面。</span></span><input type="checkbox" id="rule-recount" ${recountOn() ? 'checked' : ''} ${canManage() ? '' : 'disabled'} style="width:22px;height:22px"></label>
    </div>
    <p class="footnote">${canManage() ? '改了會同步給大家。' : '由擁有者或管理員設定。'}</p>
    ${canManage() ? pickCrossSettings() : ''}
    ${canManage() ? `<details class="steps"><summary>資料安全嗎？（公司資產）</summary>
      <ol>
        <li><b>資料放在哪：</b>只在擁有者的 Google 雲端硬碟「拍照盤點同步資料」資料夾和試算表。可以先用個人帳號，之後在「共用設定 → 搬到另一個 Google 帳號」搬到公司帳號（大家自動跟過去）。GitHub 上只有程式，沒有任何盤點資料。</li>
        <li><b>誰讀得到：</b>只有共用名單裡的人。每個人一組自己的連結碼（亂數，猜不到；雲端只存雜湊值），傳輸全程加密（HTTPS）。部署時選的「所有人」只代表可以呼叫網址，沒有連結碼一律拒絕。</li>
        <li><b>權限：</b>擁有者、管理員（可以邀請／移除人）、編輯者（可以盤點、修改）、檢視者（只能看）。改權限馬上生效，不用換連結碼。</li>
        <li><b>有人離職：</b>在「共用設定」把他「移除權限」就好，其他人不用改；他的手機下次連線時，App 裡的公司資料會自動清除。</li>
        <li><b>手機上也有一份：</b>每台裝置會存一份方便離線看；手機請設螢幕鎖。</li>
        <li><b>拍照辨識：</b>照片會送到 Google Gemini 分析。免費版的條款寫明：Google 可以用送去的內容改善產品，也可能有人工審閱。擔心的話，到 Google AI Studio 開啟付費（照用量計費），付費版不會拿去改善產品。照片裡不要拍到價格單、客戶資料。</li>
      </ol>
    </details>` : ''}
    <p class="section-title">資料</p>
    ${canManage() ? '<div class="row-actions"><button class="btn small danger" data-action="clear-all">刪除全部盤點紀錄</button></div>' : ''}
    <p class="footnote">紀錄（含照片）只存在這支手機的瀏覽器裡；要留底請用「匯出」。品項庫請到「品項 → ⋯ → 備份品項庫」。</p>
    ${myRole() === 'owner' ? '' : diagSection()}
    </div></div>
    <p class="footnote" style="margin-top:18px;text-align:center">聖佳智慧庫存 版本 ${VERSION} <span class="badge new">新</span></p>
    <div class="row-actions" style="justify-content:center"><button class="btn small secondary" data-action="force-update">檢查更新</button></div>
    <p class="footnote" style="text-align:center">有新版會自動更新；不放心就按這裡。盤點紀錄、樣品照、API Key 都不會被刪。</p>
  </main>`
}

// ───────────────────────── 品項庫畫面 ─────────────────────────
const ITEM_FILTERS = [
  { id: 'all', label: '全部', test: () => true },
  { id: 'order', label: '叫貨', test: needsOrder },
  { id: 'recount', label: '複盤', test: needsRecount },
  { id: 'diff', label: '差異', test: (it) => (diffOf(it) ?? 0) !== 0, hidden: blindMe },
  { id: 'new', label: '新的', test: (it) => it.status === 'new' },
]
const itemBadges = (it) => {
  const d = blindMe() ? null : diffOf(it)
  const rc = recountsOf(it).find(([, x]) => rcOpen(x))?.[1]
  return `${it.status === 'new' ? '<span class="badge ok">新的</span>' : ''}${needsOrder(it) ? '<span class="badge low">該叫貨</span>' : ''}${rc ? `<span class="badge low">${rc.status === 'pending' ? '待複盤' : '差異待處理'}</span>` : ''}${d ? `<span class="badge ${d < 0 ? 'bad' : 'edit'}">${d > 0 ? '+' : ''}${d}</span>` : ''}`
}
const itemRow = (it) => {
  const live = liveStock(it).map(([, st]) => st)
  const pairs = live.sort((a, b) => b.count - a.count).map((st) => [st.place, st.count])
  const search = canon([it.no, it.label, it.brand, it.model, it.spec, ...live.map((st) => st.place)].join(' '))
  return `<button class="row item-row" data-item-open="${esc(it.id)}" data-search="${esc(search)}">
    ${itemThumb(it, 44, true)}
    <span class="grow"><span class="title">${esc(itemTitle(it))}</span>${itemBadges(it)}<br><span class="meta">${esc([it.no, it.brand, it.model].filter(Boolean).join('・'))}</span>${pairs.length ? placesHtml(pairs, 4) : ''}</span>
    <span class="qty"><b>${onHand(it)}</b>${it.book != null && !blindMe() ? `<small>帳面 ${it.book}</small>` : ''}</span>${chev}</button>`
}

// ───────────────────────── 正航產品表（認識產品） ─────────────────────────
/**
 * 正航「產品存量明細表」匯進來，存成一張表（erp 裡只有一筆 catalog），同步時一次一個檔案；
 * 不會把 4000 多種一筆一筆塞進品項庫（那會讓同步變慢）。有盤到、或你按「加入品項庫」的，才變成品項。
 * products：[{ no 產品編號, name 品名規格, unit, qty 各倉合計（實際在庫量）, wh { 倉庫編號: 數量 } }]；cats：{ 類別字母: 名稱 }
 */
const ERP_ID = 'catalog'
const ERP_CATS_DEFAULT = { C: '真空邦浦' } // 在正航看到的；其他請使用者從「產品類別設定」截圖補
let erpCache
async function erpGet(force = false) {
  if (erpCache === undefined || force) erpCache = (await idb.erp.get(ERP_ID).catch(() => null)) || null
  return erpCache
}
async function erpPut(v) {
  erpCache = v
  await idb.erp.put(v)
}
/** 產品的類別：正航報表有「產品類別」欄就用它（例如 C-CC），沒有就用產品編號的第一個字母 */
const erpCat = (p) => (typeof p === 'string' ? p.trim().charAt(0).toUpperCase() : p?.cat || String(p?.no || '').trim().charAt(0).toUpperCase()) || '#'
const erpCatName = (erp, c) => erp?.cats?.[c] || ERP_CATS_DEFAULT[c] || ''
const erpCatLabel = (erp, c) => (erpCatName(erp, c) ? `${c}　${erpCatName(erp, c)}` : `${c} 類`)
/**
 * 品項庫的索引：正航產品編號 → 品項。產品總表一頁 150 列、匯入 4000 多種，一個一個從頭找太慢，先建好對照表。
 * 先看「接上的正航料號」（erpNo）；沒有才看料號一樣的，而且那一筆不能已經接上別的正航產品
 * （不然一個品項會同時算成兩個正航產品）。
 */
function erpIndex(items) {
  const byErp = new Map()
  const byNo = new Map()
  for (const it of items) {
    if (it.erpNo && !byErp.has(it.erpNo)) byErp.set(it.erpNo, it)
    const k = canon(it.no)
    if (k && !byNo.has(k)) byNo.set(k, it)
  }
  return {
    of: (no) => {
      const linked = byErp.get(no)
      if (linked) return linked
      const same = byNo.get(canon(no))
      return same && (!same.erpNo || same.erpNo === no) ? same : null
    },
    link: (it, no) => byErp.set(no, it),
  }
}
/** 這個產品在品項庫裡的那一筆（料號＝產品編號）；只找一個時用，找很多個請先 erpIndex() */
const erpItemOf = (items, no) => erpIndex(items).of(no)
/** 讀 CSV／Excel 貼上的表格：逗號或 Tab 分隔，欄位可以用引號包（品名裡有逗號也沒問題） */
function parseDelimited(text) {
  const s = String(text).replace(/^﻿/, '')
  const firstLine = s.split(/\r?\n/)[0] || ''
  const sep = firstLine.includes('\t') ? '\t' : ','
  const rows = []
  let row = []
  let cell = ''
  let q = false
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]
    if (q) {
      if (ch === '"' && s[i + 1] === '"') {
        cell += '"'
        i++
      } else if (ch === '"') q = false
      else cell += ch
    } else if (ch === '"') q = true
    else if (ch === sep) {
      row.push(cell)
      cell = ''
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && s[i + 1] === '\n') i++
      row.push(cell)
      rows.push(row)
      row = []
      cell = ''
    } else cell += ch
  }
  if (cell || row.length) {
    row.push(cell)
    rows.push(row)
  }
  return rows.filter((r) => r.some((c) => c.trim()))
}
/**
 * 正航的報表 → 產品清單。認得兩種：
 * - 產品存量明細表：產品編號、品名規格、單位、倉庫編號、實際在庫量（同一個產品在幾個倉庫就有幾列，這裡合併）
 * - 歷史庫存一覽表：多了「產品類別」「類別名稱」（標題列不一定在第一列，上面有公司名稱、報表名稱）
 * 成本、售價那些欄位一律不讀。回傳 { products, cats }。
 */
function parseErp(text) {
  const all = parseDelimited(text)
  const hi = all.findIndex((r) => r.some((h) => h.replace(/\s/g, '').includes('產品編號')))
  if (hi < 0) throw new Error('這不是正航的產品報表：要有「產品編號」這一欄（產品存量明細表、歷史庫存一覽表都可以）')
  const rows = all.slice(hi)
  const head = rows[0].map((h) => h.replace(/\s/g, ''))
  const col = (...ws) => head.findIndex((h) => ws.some((w) => h.includes(w)))
  const c = { no: col('產品編號'), name: col('品名規格', '品名'), unit: col('計量單位', '單位'), wh: col('倉庫編號'), qty: col('實際在庫'), onhand: col('現有庫存', '現有數量'), cat: col('產品類別'), catName: col('類別名稱') }
  if (c.no < 0 || c.name < 0) throw new Error('這不是正航的產品報表：要有「產品編號」和「品名規格」')
  const qcol = c.qty >= 0 ? c.qty : c.onhand
  const num = (v) => {
    const n = Number(String(v ?? '').replace(/[,\s]/g, ''))
    return Number.isFinite(n) ? n : 0
  }
  const map = new Map()
  const cats = {}
  for (const r of rows.slice(1)) {
    const no = String(r[c.no] || '').trim()
    if (!no) continue
    const p = map.get(no) || { no, name: String(r[c.name] || '').trim(), unit: c.unit >= 0 ? String(r[c.unit] || '').trim() : '', qty: 0, wh: {} }
    const cat = c.cat >= 0 ? String(r[c.cat] || '').trim() : ''
    if (cat) {
      p.cat = cat
      const cn = c.catName >= 0 ? String(r[c.catName] || '').trim() : ''
      if (cn) cats[cat] = cn
    }
    const w = c.wh >= 0 ? String(r[c.wh] || '').trim() : ''
    const q = qcol >= 0 ? num(r[qcol]) : 0
    p.qty += q
    if (w) p.wh[w] = (p.wh[w] || 0) + q
    map.set(no, p)
  }
  if (!map.size) throw new Error('沒有讀到任何產品編號')
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  return { products: [...map.values()].sort((a, b) => cmp(a.no, b.no)), cats }
}
/** 讀檔：Excel 另存的 CSV 可能是 UTF-8 或 Big5（Windows 預設），兩種都認 */
async function readTextFile(file) {
  const buf = await file.arrayBuffer()
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buf)
  } catch {
    return new TextDecoder('big5').decode(buf)
  }
}
/**
 * 匯入：存成一張表；品項庫裡料號對得上的，帳面改成正航的數量。
 * 點貨交叉比對的基準（4.7）：這次報表裡有的產品，記下 at＝報表的時間（qty 就是這次正航的數量）＝基準 { qty, at }；
 * 這次沒出現的產品照舊（基準不變）。每次都記；交叉比對開關打開才比、才顯示。
 * reportAt：報表的時間——選檔匯入用檔案的修改時間（比匯入 App 的時間更接近正航匯出的時間）；貼上的用現在。
 */
async function importErp(text, { reportAt = 0 } = {}) {
  const parsed = parseErp(text)
  const old = await erpGet()
  const now = Date.now()
  // 檔案時間不合理（未來、比上次基準還早一年以上）就用現在
  const at = reportAt > 0 && reportAt <= now && reportAt > now - 365 * 86400000 ? reportAt : now
  // 兩種報表可以輪流匯：這次沒有的欄位（類別、別的產品）用上次的補
  const oldBy = new Map((old?.products || []).map((p) => [p.no, p]))
  const products = parsed.products.map((p) => ({ ...(p.cat || !oldBy.get(p.no)?.cat ? p : { ...p, cat: oldBy.get(p.no).cat }), at }))
  // 交叉比對（測試版，只有開關打開的擁有者／管理員）：用上次的基準和這次的數量，比上次匯入到這張報表之間的點貨紀錄
  const cross = crossOn() ? { at, ...crossCheck(old?.products || [], parsed.products, await idb.picks.all().catch(() => []), at) } : null
  const seen = new Set(products.map((p) => p.no))
  for (const p of old?.products || []) if (!seen.has(p.no)) products.push(p)
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  products.sort((a, b) => cmp(a.no, b.no))
  await erpPut({ id: ERP_ID, createdAt: old?.createdAt || now, at: now, products, cats: { ...(old?.cats || {}), ...parsed.cats } })
  const items = await itemsAll(true)
  const index = erpIndex(items)
  let booked = 0
  for (const p of products) {
    const it = index.of(p.no)
    if (!it) continue
    if (it.erpNo !== p.no) {
      it.erpNo = p.no
      it.erpAt = Date.now()
    }
    index.link(it, p.no)
    if (it.book !== p.qty) {
      it.book = p.qty
      ;(it.moves ||= []).push({ at: Date.now(), kind: 'set', qty: p.qty, from: 'erp' }) // 很舊的品項沒有進出紀錄欄位
      booked++
    }
    await putItem(it)
  }
  return { products: products.length, stocked: products.filter((p) => p.qty > 0).length, booked, cats: Object.keys(parsed.cats).length, cross }
}
/** 從產品表把一個產品記進品項庫（料號＝產品編號、帳面＝正航數量） */
async function erpLink(no, intoId = '') {
  const erp = await erpGet()
  const p = erp?.products.find((x) => x.no === no)
  if (!p) return null
  const items = await itemsAll()
  let it = erpItemOf(items, no)
  // 合併到已經有的品項（例如盤點時自動建立的 P0001）：接上正航料號、帳面改成正航的數量，不另外新增一筆
  const into = !it && intoId ? items.find((x) => x.id === intoId) : null
  if (into) {
    // 合併前的樣子：給「復原」用
    const prev = { erpNo: into.erpNo || '', unit: into.unit, book: into.book ?? null }
    into.erpNo = p.no
    into.erpAt = Date.now()
    if (!into.unit && p.unit) into.unit = p.unit
    if (into.book !== p.qty) {
      into.book = p.qty
      ;(into.moves ||= []).push({ at: Date.now(), kind: 'set', qty: p.qty, from: 'erp' })
    }
    await putItem(into)
    return { item: into, prev }
  }
  if (!it) {
    it = newItem(items, { label: p.name, brand: '', model: '', spec: '' }, { no: p.no, erpNo: p.no, erpAt: Date.now(), unit: p.unit, status: 'ok', book: p.qty, moves: [{ at: Date.now(), kind: 'set', qty: p.qty, from: 'erp' }] })
    items.unshift(it)
  } else if (it.erpNo !== p.no) {
    it.erpNo = p.no
    it.erpAt = Date.now()
  }
  await putItem(it)
  return { item: it }
}
/**
 * 復原「合併到已經有的品項」：拿掉正航料號、帳面改回原本的。
 * 帳面用「再記一筆設定」改回去（不是刪掉紀錄）：進出紀錄在多台之間是合起來的，刪掉的會被別台補回來。
 */
async function erpUnlink(id, prev, no) {
  const it = (await itemsAll()).find((x) => x.id === id)
  if (!it || it.erpNo !== no) return false
  it.erpNo = prev.erpNo
  it.erpAt = Date.now() // 比合併那次新：同步時以「取消」為準
  if (prev.unit === undefined) delete it.unit
  else it.unit = prev.unit
  if (it.book !== prev.book) {
    ;(it.moves ||= []).push({ at: Date.now(), kind: 'set', qty: prev.book, from: 'undo', why: `取消接上正航 ${no}` })
    it.book = prev.book
  }
  await putItem(it)
  return true
}
const erpSearchKey = (p) => canon(`${p.no ?? ''} ${p.name ?? ''}`)
/** 只留英文和數字（型號比對用）：「DML 083S」→「dml083s」 */
const alnum = (s) => canon(s).replace(/[^a-z0-9]/g, '')
/** 比對時可以忽略的分隔符號（跟 canon 拿掉的一樣） */
const TOKEN_SEP = `[\\s\\-_/.,，、・·()（）\\[\\]【】"'“”‘’]*`
const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
/**
 * needle 有沒有「完整」出現在 hay 裡：中間的空白、橫線不計（DML083＝DML 083），
 * 但前後不能再接英文或數字——「KP 1」不會對到「KP15」、「DML 083」不會對到「DML083S」（S＝焊接，是另一種產品）。
 */
function hasToken(hay, needle, tail = true) {
  const re = tokenRe(needle, tail)
  return !!re && re.test(normHay(hay))
}
/** 比對前先把要找的那段字整理好（全形→半形、大寫→小寫）；點貨搜尋 4000 多種產品時每種只整理一次 */
const normHay = (s) => String(s ?? '').normalize('NFKC').toLowerCase()
/**
 * 「完整出現」的比對規則（hasToken 用）：同一個詞只建一次（點貨搜尋每打一個字要比 4000 多種產品）。
 * tail＝false：後面可以再接英數字＝「開頭一樣」（打到一半時列「開頭一樣的」用，不能拿來自動配對）。
 */
const TOKEN_RE = new Map()
/** 有數字的詞：小數點、斜線要照寫（2.5L 不是 25L、3/8 不是 38），其他分隔符號照樣不計 */
const TOKEN_SEP_NUM = `[\\s\\-_,，、・·()（）\\[\\]【】"'“”‘’]*`
function tokenKey(needle) {
  const s = normHay(needle)
  if (!/\d/.test(s)) return canon(needle)
  return s.replace(/[\s\-_,，、・·()（）[\]【】"'“”‘’]/g, '').replace(/^[./]+|[./]+$/g, '')
}
function tokenRe(needle, tail = true) {
  const c = tokenKey(needle)
  if (!c) return null
  const key = `${tail ? 1 : 0}${c}`
  let re = TOKEN_RE.get(key)
  if (!re) {
    const num = /\d/.test(c)
    const body = [...c].map(reEsc).join(num ? TOKEN_SEP_NUM : TOKEN_SEP)
    // 前後不能再接英數字；數字開頭、結尾的也不能接「數字＋小數點或斜線」：5L 不會對到 2.5L、8 不會對到 3/8、2 不會對到 2.5
    const head = /^[0-9]/.test(c) ? '(?:^|[^a-z0-9./]|[^0-9][./])' : /^[a-z]/.test(c) ? '(?:^|[^a-z0-9])' : ''
    const end = !tail ? '' : /[0-9]$/.test(c) ? '(?![a-z0-9]|[./][0-9])' : /[a-z]$/.test(c) ? '(?![a-z0-9])' : ''
    re = new RegExp(head + body + end)
    if (TOKEN_RE.size > 2000) TOKEN_RE.clear()
    TOKEN_RE.set(key, re)
  }
  return re
}
/**
 * 品項庫裡可能就是這個正航產品的（還沒接上任何正航產品的）。給「要合併到 P0001 嗎？」用，最多 3 個：
 * 1. 品項有型號：型號要完整出現在正航品名裡；對不到＝不同產品，品名一樣也不建議。
 * 2. 沒有型號：品名要出現；有規格的規格也要出現。
 * 3. 只對到品名（沒型號、沒規格）：品名要 3 個字以上（「銅管」「三通」幾乎每個品名都有），排最後、標「只對到品名」。
 * erpNos：正航全部的產品編號（canon 過）；料號跟某個正航產品一樣的，已經算是那個產品的，不能再接別的。
 */
function erpSimilar(items, p, erpNos = new Set()) {
  const score = (it) => {
    if (alnum(it.model).length >= 3) return hasToken(p.name, it.model) ? 3 : 0
    if (!it.label || !hasToken(p.name, it.label)) return 0
    if (it.spec) return hasToken(p.name, it.spec) ? 2 : 0
    return canon(it.label).length >= 3 ? 1 : 0
  }
  return items
    .filter((it) => !it.erpNo && !erpNos.has(canon(it.no)))
    .map((it) => ({ it, score: score(it) }))
    .filter((x) => x.score)
    .sort((a, b) => b.score - a.score)
    .slice(0, 3)
    .map(({ it, score }) => ({ it, nameOnly: score === 1 }))
}
/** 認識產品：先看類別、再看那一類的東西；可以搜尋編號或名字 */
async function viewCatalog() {
  const erp = await erpGet()
  const st = (state.erp ??= { cat: null, q: '', stocked: true, more: 0 })
  const items = await itemsAll()
  const back = backBtn('items', '品項')
  if (!erp)
    return `
  <main class="app">
    <div class="nav">${back}</div>
    <h1 class="large-title">產品總表</h1>
    <p class="subtitle">把正航的產品表匯進來：每一種產品的庫存、放在哪裡、照片，點貨對單用。</p>
    <div class="hint-card stack">
      <div><b>還沒匯入正航產品表。</b>在公司電腦的正航：報表 → 庫存管理 → 存貨狀況報表 → <b>產品存量明細表</b> → 確定 → 存成 Excel；再另存成 CSV 檔，或整張複製貼上。</div>
      ${canManage() ? '<button class="btn small" data-action="erp-import">匯入正航產品表</button>' : '<div class="muted">匯入正航產品表會改大家的帳面數：請擁有者或管理員匯入。</div>'}
    </div>
  </main>`
  const all = erp.products
  const stockedN = all.filter((p) => p.qty > 0).length
  const pool = st.stocked ? all.filter((p) => p.qty > 0) : all
  const cats = new Map()
  for (const p of pool) cats.set(erpCat(p), (cats.get(erpCat(p)) || 0) + 1)
  const catList = [...cats].sort((a, b) => b[1] - a[1])
  return `
  <main class="app">
    <div class="nav">${back}${canManage() ? '<span class="nav-right"><button class="btn small plain" data-action="erp-cats">類別名稱</button><button class="btn small plain" data-action="erp-import">重新匯入</button></span>' : ''}</div>
    <h1 class="large-title">產品總表</h1>
    <p class="subtitle">正航 ${all.length} 種・有庫存 ${stockedN} 種・${fmtTime(erp.at)} 匯入。點類別或直接搜尋編號、品名。</p>
    <input class="field search" id="erp-q" type="search" placeholder="搜尋產品編號、品名（全部 ${all.length} 種）" autocomplete="off" enterkeyhint="search" value="${esc(st.q)}">
    <div class="chips"><button class="chip ${st.stocked ? 'on' : ''}" data-action="erp-stocked" aria-pressed="${st.stocked}">${st.stocked ? icon('check', 18) : ''}只看有庫存的 <small>${stockedN}</small></button>${st.cat ? `<button class="chip" data-erp-cat="">${icon('chevron-left', 18)}所有類別</button>` : ''}</div>
    <div id="erp-list">${erpListHtml(erp, items, catList)}</div>
  </main>`
}
/** 清單那一塊（搜尋時只重畫這裡，打字的框不會跳掉） */
function erpListHtml(erp, items, catList) {
  const st = state.erp
  const pool = st.stocked ? erp.products.filter((p) => p.qty > 0) : erp.products
  const q = canon(st.q)
  const PAGE = 150
  const index = erpIndex(items) // 每打一個字都會重畫：先建對照表，不要每一列都從頭找
  const rowOf = (p) => {
    const it = index.of(p.no)
    const places = it ? liveStock(it).map(([, s]) => [s.place, null]).filter(([pl]) => pl) : []
    const meta = [p.no, p.unit, blindMe() ? '' : `庫存 ${p.qty ?? 0}`].filter(Boolean).join('・')
    return `<button class="row" data-erp-no="${esc(p.no)}">${it ? itemThumb(it, 40, true) : `<span class="thumb ph" style="width:40px;height:40px" aria-hidden="true">${esc(erpCat(p))}</span>`}<span class="grow"><span class="title">${esc(p.name)}</span><br><span class="meta">${esc(meta)}</span>${places.length ? placesHtml(places, 3) : ''}</span>${it ? '<span class="badge ok">品項</span>' : ''}${chev}</button>`
  }
  const list = (arr, label) => {
    const shown = arr.slice(0, PAGE + st.more)
    return `<section class="item-sec"><p class="section-title">${esc(label)}（${arr.length}）</p><div class="group">${shown.map(rowOf).join('') || '<div class="row muted">沒有</div>'}</div>${arr.length > shown.length ? `<button class="btn secondary block" data-action="erp-more" style="margin-top:10px">再顯示 ${Math.min(PAGE, arr.length - shown.length)} 種（還有 ${arr.length - shown.length}）</button>` : ''}</section>`
  }
  if (q) return list(pool.filter((p) => erpSearchKey(p).includes(q)), `找到`)
  if (st.cat) return list(pool.filter((p) => erpCat(p) === st.cat), erpCatLabel(erp, st.cat))
  if (!catList) {
    const cats = new Map()
    for (const p of pool) cats.set(erpCat(p), (cats.get(erpCat(p)) || 0) + 1)
    catList = [...cats].sort((a, b) => b[1] - a[1])
  }
  const unnamed = catList.filter(([c]) => !erpCatName(erp, c)).length
  return `<div class="group">${catList.map(([c, n]) => `<button class="row" data-erp-cat="${esc(c)}"><span class="thumb ph" style="width:40px;height:40px" aria-hidden="true">${esc(c)}</span><span class="grow"><span class="title">${esc(erpCatLabel(erp, c))}</span><br><span class="meta">${n} 種</span></span>${chev}</button>`).join('')}</div>${unnamed ? `<p class="footnote">${unnamed} 個類別只有字母、還沒有名稱：到正航「產品類別設定」看代號對應什麼，${canManage() ? '按右上「類別名稱」填進來' : '請擁有者或管理員填類別名稱'}。</p>` : ''}`
}
/** 一個產品：編號、名稱、單位、各倉庫數量、放哪裡、照片；可以加入品項庫 */
async function erpProductSheet(no) {
  const erp = await erpGet()
  const p = erp?.products.find((x) => x.no === no)
  if (!p) return toast('找不到這個產品')
  const items = await itemsAll()
  const it = erpItemOf(items, no)
  const stock = it ? liveStock(it) : []
  sheet(
    `<h2 class="sheet-title">${esc(p.name)}</h2>
     <p class="sheet-sub">${esc(p.no)}・${esc(erpCatLabel(erp, erpCat(p)))}${p.unit ? `・單位：${esc(p.unit)}` : ''}</p>
     ${it?.photo ? zoomWrap(`<img class="sheet-photo" src="${itemUrl(it)}" alt="${esc(p.name)}的照片">`, itemZoom(it), 'block lg') : ''}
     <div class="group" style="margin-top:10px">
       ${blindMe() ? '' : `<div class="row"><span class="grow"><span class="title">正航庫存</span><br><span class="meta">${Object.entries(p.wh || {}).map(([w, n]) => `倉庫 ${esc(w)}：${esc(n)}`).join('・') || '沒有倉庫資料'}</span></span><b>${esc(p.qty ?? 0)}</b></div>`}
       <div class="row"><span class="grow"><span class="title">放在哪裡</span><br>${stock.length ? placesHtml(stock.map(([, s]) => [s.place, s.count])) : `<span class="meta">${it ? '還沒盤點到' : '還沒記進品項庫，盤點到才知道'}</span>`}</span></div>
     </div>
     ${it ? `<button class="btn block" id="e-open" style="margin-top:12px">打開品項（照片、盤點紀錄）</button>` : `<button class="btn block edit-only" id="e-link" style="margin-top:12px">加入品項庫</button><div id="e-dup" hidden></div><p class="footnote">加入後：盤點到會記位置、可以放樣品照、看差異。</p>`}`,
    (el, close) => {
      el.querySelector('#e-open')?.addEventListener('click', () => {
        close()
        state.itemId = it.id
        go('item')
      })
      const link = async (intoId = '') => {
        const made = await erpLink(no, intoId)
        close()
        if (!made) return toast('加不進去，請再試一次')
        const { item, prev } = made
        state.itemId = item.id
        go('item')
        if (!prev) return toast(`已記進品項庫：${item.label}`)
        // 盲盤：盤點的人看不到帳面數，提示裡也不寫數字
        const bookTxt = blindMe() || prev.book === item.book ? '' : prev.book == null ? `，帳面改成 ${item.book}（原本沒有）` : `，帳面 ${prev.book} → ${item.book}`
        toast(`已合併：${item.no} 接上正航 ${no}${bookTxt}`, {
          label: '復原',
          run: async () => {
            const ok = await erpUnlink(item.id, prev, no)
            toast(ok ? `已復原：${item.no} 沒有接上正航 ${no}` : '這一筆已經改過了，沒辦法復原')
            if (['item', 'catalog', 'items'].includes(state.view)) render()
          },
        })
      }
      el.querySelector('#e-link')?.addEventListener('click', () => {
        // 加入前先找品項庫裡是不是已經有同一個（例如盤點時自動建立的）：有的話問要不要合併，免得變成兩筆、盤差看錯
        const similar = erpSimilar(items, p, new Set(erp.products.map((x) => canon(x.no))))
        if (!similar.length) return link()
        const blind = blindMe()
        const box = el.querySelector('#e-dup')
        box.hidden = false
        box.innerHTML = `<p class="section-title" style="margin-top:14px">品項庫裡可能已經有了</p>
          <div class="group">${similar.map(({ it: x, nameOnly }) => `<button class="row" data-into="${esc(x.id)}">${itemThumb(x, 40, true)}<span class="grow"><span class="title">合併到 ${esc(x.no)}</span>${nameOnly ? '<span class="badge low">只對到品名</span>' : ''}<br><span class="meta">${esc(itemTitle(x))}${x.model ? `・${esc(x.model)}` : ''}・實盤 ${onHand(x)}</span></span>${chev}</button>`).join('')}</div>
          <p class="footnote">合併：正航料號 ${esc(no)} 接到那一筆，${blind ? '帳面數改成正航的數量' : `帳面數改成正航的 ${esc(p.qty)}`}。${similar.some((x) => x.nameOnly) ? '「只對到品名」的不一定是同一個產品，請看清楚型號、規格。' : ''}</p>
          <button class="btn plain block" id="e-new" style="margin-top:8px">不一樣，還是新增一筆</button>`
        el.querySelector('#e-link').hidden = true
        glueTails(box)
        box.querySelectorAll('[data-into]').forEach(
          (b) =>
            (b.onclick = () => {
              const x = similar.find((s) => s.it.id === b.dataset.into)?.it
              if (!x) return
              const change = blind ? '・帳面數改成正航的數量（盲盤中不顯示數字）' : x.book === p.qty ? `・帳面不變（都是 ${p.qty}）` : x.book == null ? `・帳面改成 ${p.qty}（原本沒有）` : `・帳面 ${x.book} → ${p.qty}`
              if (!confirm(`合併到 ${x.no}「${itemTitle(x)}」？\n\n・接上正航 ${no}「${p.name}」\n${change}\n\n合併後幾秒內可以按「復原」。`)) return
              link(x.id)
            }),
        )
        box.querySelector('#e-new').onclick = () => link()
      })
    },
  )
}
/** 類別名稱：一行一個「字母=名稱」，跟著產品表一起同步 */
async function erpCatsSheet() {
  const erp = await erpGet()
  if (!erp) return toast('先匯入正航產品表')
  const letters = [...new Set(erp.products.map((p) => erpCat(p)))].sort()
  const text = letters.map((c) => `${c}=${erpCatName(erp, c)}`).join('\n')
  sheet(
    `<h2 class="sheet-title">類別名稱</h2>
     <p class="sheet-sub">產品編號開頭的字母是正航的「產品類別」。到正航「系統 → 共用資料 → 產品資料設定 → 產品類別設定」看代號對應的名稱，填在等號後面（例如 C=真空邦浦）。</p>
     <textarea class="field" id="ec-text" rows="${Math.min(14, letters.length + 1)}">${esc(text)}</textarea>
     <button class="btn block" id="ec-save" style="margin-top:12px">儲存</button>`,
    (el, close) => {
      el.querySelector('#ec-save').onclick = async () => {
        const cats = {}
        for (const line of el.querySelector('#ec-text').value.split('\n')) {
          const m = /^\s*([A-Za-z0-9#])\s*[=＝:：]\s*(.*?)\s*$/.exec(line)
          if (m && m[2]) cats[m[1].toUpperCase()] = m[2]
        }
        await erpPut({ ...erp, cats })
        close()
        toast('類別名稱存好了')
        render()
      }
    },
  )
}
/** 匯入正航產品表：選 CSV 檔，或整張複製貼上 */
function erpImportSheet() {
  sheet(
    `<h2 class="sheet-title">匯入正航產品表</h2>
     <ol class="steps-list">
       <li>公司電腦的正航：報表 → 庫存管理 → 存貨狀況報表 → <b>產品存量明細表</b>（全部產品）或<b>歷史庫存一覽表</b>（有類別名稱）→ 確定 → 存成 Excel。兩種都匯，資料會合併。</li>
       <li>在 Excel 按「另存新檔」，存檔類型選 <b>CSV</b>（UTF-8 或一般的都可以），傳到這台。</li>
       <li>按下面「選 CSV 檔」。電腦上也可以直接把整張表（含標題列）複製，貼在下面。</li>
     </ol>
     <label class="btn block secondary" style="margin-top:10px">選 CSV 檔<input type="file" accept=".csv,.txt,.tsv,text/csv,text/plain" id="ei-file" class="sr-only"></label>
     <textarea class="field" id="ei-text" rows="4" placeholder="產品編號&#9;品名規格&#9;單位&#9;倉庫編號&#9;…&#9;實際在庫量&#10;C-043&#9;真空幫浦 …&#9;台&#9;02&#9;…&#9;2" style="margin-top:10px"></textarea>
     <button class="btn block" id="ei-go" style="margin-top:12px">匯入貼上的內容</button>
     <p class="footnote">同一個產品在幾個倉庫會合併成一筆。品項庫裡料號對得上的，帳面會改成正航的數量；安全庫存、成本、售價這裡都不會讀。</p>`,
    (el, close) => {
      const run = async (text, reportAt = 0) => {
        try {
          // 交叉比對開著：先同步一次，拿到別台剛點完的點貨紀錄再比（同步失敗就照常比，結果頁會寫最後同步的時間）
          if (crossOn() && syncReady()) {
            toast('先同步一次，拿別台的點貨紀錄…')
            await syncNow().catch(() => {})
          }
          const r = await importErp(text, { reportAt })
          close()
          const msg = `匯入完成：${r.products} 種（有庫存 ${r.stocked} 種）${r.cats ? `、${r.cats} 個類別名稱` : ''}${r.booked ? `，更新 ${r.booked} 個品項的帳面` : ''}`
          state.erp = { cat: null, q: '', stocked: true, more: 0 }
          go('catalog')
          // 交叉比對開關關著：跟以前一樣，只有一行提示；開著：存下這次的比對結果，跳出「這次有 N 項要查一下」
          if (!r.cross) return toast(msg)
          // 這次沒有可以比的（上次匯入後沒人點貨）：不要蓋掉上次的結果
          if (r.cross.first || r.cross.checked) ls.set(LS.pickCrossLast, JSON.stringify(r.cross))
          crossDoneSheet(msg, r.cross)
        } catch (e) {
          toast(e.message)
        }
      }
      el.querySelector('#ei-file').onchange = async (e) => {
        const f = e.target.files?.[0]
        if (f) run(await readTextFile(f), f.lastModified || 0)
      }
      el.querySelector('#ei-go').onclick = () => run(el.querySelector('#ei-text').value)
    },
  )
}

/** 儲位數：建好的儲位＋品項裡用到、但還沒建的代號（不然會寫「0 個儲位」卻看到 A-01） */
const placeCount = (items) => new Set([...locations().map((l) => canon(l.code)), ...items.flatMap((it) => liveStock(it).map(([, st]) => canon(st.place)).filter(Boolean))]).size
async function viewItems() {
  const items = await itemsAll(true)
  const erp = await erpGet()
  const filters = ITEM_FILTERS.filter((x) => !x.hidden?.() && (x.id !== 'recount' || items.some(needsRecount)))
  const f = filters.find((x) => x.id === state.itemFilter) || filters[0]
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  const list = items.filter(f.test).sort((a, b) => cmp(a.label, b.label) || cmp(a.spec, b.spec) || cmp(a.model, b.model))
  const byLabel = new Map()
  for (const it of list) byLabel.set(it.label, [...(byLabel.get(it.label) || []), it])
  return `
  <main class="app">
    <div class="nav"><button class="btn small plain" data-go="locations">${icon('shelf', 18)}儲位</button><span class="nav-right"><button class="icon-btn" data-action="items-more" aria-label="匯入、匯出、備份" title="匯入、匯出、備份">${icon('more', 24)}</button><button class="btn small edit-only" data-action="item-add">${icon('plus', 18)}新增</button></span></div>
    <h1 class="large-title">品項庫</h1>
    <p class="subtitle">${items.length ? `${items.length} 種商品・${locations().length || !placeCount(items) ? `${placeCount(items)} 個儲位` : `用到 ${placeCount(items)} 個位置（還沒建立儲位）`}。盤點按「完成」就會自動更新。` : '盤點按「完成」，數到的東西就會自動記進來。'}</p>
    <button class="report-card" data-go="catalog" style="margin-bottom:12px"><span class="row-ico" aria-hidden="true">${icon('tag')}</span><span class="grow"><span class="report-card-kicker">產品總表（正航）</span><span class="report-card-nums">${erp ? `<b>${erp.products.length}</b> 種・有庫存 <b>${erp.products.filter((p) => p.qty > 0).length}</b> 種` : '匯入正航產品表'}</span><span class="meta">${erp ? '查庫存、放在哪裡、照片；點貨對單用' : '正航的全部產品：查庫存、放哪裡，點貨對單不會錯'}</span></span>${chev}</button>
    ${
      items.length
        ? `<div class="seg" role="tablist" aria-label="篩選">${filters.map((x) => `<button role="tab" aria-selected="${x.id === f.id}" data-item-filter="${x.id}">${x.label} ${items.filter(x.test).length}</button>`).join('')}</div>
           ${f.id === 'recount' ? `<p class="footnote" style="margin:4px 2px 10px">這些品項盤到的數量跟帳面不一樣。請<b>另一個人</b>到那一格再數一次，點進去按「我來複盤」；複盤時看不到第一次的數字，才不會受影響。</p>` : ''}
           ${f.id === 'all' && items.some((it) => it.status === 'new') ? `<button class="sum-doubt tip edit-only" data-item-filter="new"><span class="sum-dot" aria-hidden="true">!</span><span class="grow"><b>${items.filter((it) => it.status === 'new').length} 個新的品項，請確認名稱</b><br><span class="meta">AI 盤點時建立的：名稱對就確認，重複的就合併</span></span>${chev}</button>` : ''}
           ${f.id === 'order' && list.length ? `<button class="btn secondary block" data-action="order-copy" style="margin-bottom:6px">複製叫貨清單（貼到 LINE）</button>` : ''}
           ${f.id === 'order' && !list.length ? '' : '<input class="field search" id="item-search" type="search" placeholder="搜尋品名、型號、料號、儲位" autocomplete="off" enterkeyhint="search">'}
           ${
             list.length
               ? `<div class="item-secs${byLabel.size >= 4 && list.length >= 8 ? ' multi' : ''}">${[...byLabel]
                   .map(([label, its]) => `<section class="item-sec"><p class="section-title">${esc(label)}（${its.length}）</p><div class="group${byLabel.size >= 4 && list.length >= 8 ? '' : its.length >= 2 ? ' cols-2' : ''}">${its.map(itemRow).join('')}</div></section>`)
                   .join('')}</div>`
               : f.id === 'order'
                 ? `<div class="hint-card stack" style="margin-top:12px">
                      <div><b>還沒有要叫貨的。</b></div>
                      <div>先告訴 App 每一種「<b>剩幾個就要叫貨</b>」：數量剩這麼多（或更少）時，就會出現在這裡。${items.some((it) => it.safety != null) ? '' : '<br>目前每一種都還沒設定。'}</div>
                      <button class="btn small edit-only" data-action="safety-pick">設定「剩幾個就要叫貨」（安全庫存）</button><button class="btn small secondary edit-only" data-action="import">很多種一次設：貼上 Excel</button>
                    </div>`
                 : `<div class="empty"><p>${f.id === 'diff' ? '實盤跟帳面都一樣。<br>（要先在品項裡設定「帳面數」才會比對）' : '沒有新的品項，都確認過了。'}</p></div>`
           }
           <p class="empty" id="search-empty" hidden>找不到。可以到「查型號」用型號找替代品。</p>
           <details class="steps" style="margin-top:18px"><summary>品項庫怎麼用？</summary>
             <ol>
               <li><b>記進來：</b>盤點完按「完成・記進品項庫」，數到的東西會自動記進來。也可以按右上「＋ 新增」，或到「查型號」拍標籤加入。</li>
               <li><b>確認名稱：</b>標「新的」是 AI 自動建立的，點進去看名稱對不對，對就按「確認」。</li>
               <li><b>叫貨提醒（安全庫存）：</b>點進一個品項，設「剩幾個就要叫貨」；數量少於這個數字，就會出現在上面的「叫貨」。<b>很多種一次設：</b>右上「⋯」→「貼上 Excel 清單」，表格有「安全庫存」這一欄就會一起設好，不用一個一個打。</li>
               <li><b>帳面數：</b>點進品項按「設定帳面數 → 用實盤數」當起點；之後進貨按「＋ 進貨」、賣掉按「－ 賣出」。下次盤點數量跟帳面不一樣，就會出現在「差異」。</li>
               <li><b>複盤：</b>盤到的跟帳面不一樣，會出現在「複盤」：請另一個人再數一次，兩次一樣才算確定，再選原因；擁有者或管理員決定要不要調整帳面。</li>
             </ol>
           </details>`
        : `<div class="hint-card stack">
             <div><b>品項庫會自己長出來</b>：不用先建好。每次盤點按「完成・記進品項庫」，數到的每一種都會變成一筆（有料號、在哪裡、幾個）。</div>
             <div>也可以：</div>
             <div class="row-actions"><button class="btn small" data-go="lookup">${icon('camera', 20)}拍型號加進來</button><button class="btn small secondary" data-action="import">貼上 Excel 清單</button><button class="btn small secondary" data-go="locations">建立儲位</button></div>
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
        ? `<div class="hint-card stack edit-only"><div><b>盤點時自動建立的。</b>名稱、尺寸對嗎？跟別的品項重複就合併。</div><div class="row-actions"><button class="btn small" data-action="item-confirm">${icon('check', 18)}對，確認</button><button class="btn small secondary" data-action="item-edit">修改</button><button class="btn small secondary" data-action="item-merge">合併到…</button></div></div>`
        : ''
    }
    <div class="detail-cols"><div class="detail-main">
    <section class="summary">
      ${
        blindMe()
          ? `<div class="stock-nums"><div><span class="stock-label">實盤</span><span class="stock-big">${total}</span><span class="stock-sub">${last ? `最近 ${fmtTime(last)}` : '還沒盤過'}</span></div></div>
             <p class="footnote" style="margin:10px 0 0">盲盤中：帳面數和差異只有擁有者、管理員看得到。這樣盤點時只會照實際數，不會被帳面數影響。</p>`
          : `<div class="stock-nums">
        <div><span class="stock-label">實盤</span><span class="stock-big">${total}</span><span class="stock-sub">${last ? `最近 ${fmtTime(last)}` : '還沒盤過'}</span></div>
        <div><span class="stock-label">帳面</span><span class="stock-big">${it.book ?? '—'}</span><span class="stock-sub">${it.book == null ? '還沒設定' : '進貨加、賣出減'}</span></div>
        <div><span class="stock-label">差異</span><span class="stock-big ${d < 0 ? 'neg' : d > 0 ? 'pos' : ''}">${d == null ? '—' : `${d > 0 ? '+' : ''}${d}`}</span><span class="stock-sub">${d == null ? '設定帳面數才比' : d > 0 ? '盤盈（多了）' : d < 0 ? '盤虧（少了）' : '一樣'}</span></div>
      </div>`
      }
      <div class="row-actions edit-only" style="margin-top:14px"><button class="btn small secondary" data-action="move-in">＋ 進貨</button><button class="btn small secondary" data-action="move-out">－ 賣出</button>${blindMe() ? '' : '<button class="btn small secondary" data-action="book-set">設定帳面數</button>'}</div>
    </section>
    ${recountSection(it)}
    <p class="section-title">在哪裡（${stock.length} 個位置）</p>
    ${
      stock.length
        ? `<div class="group">${stock.map(([k, st]) => `<div class="row stock-row"><span class="grow"><span class="title">${placeHtml(st.place)}</span><span class="meta">${fmtTime(st.at)} 盤點</span></span><span class="qty"><b>${st.count}</b></span><button class="icon-btn small edit-only" data-stock-del="${esc(k)}" aria-label="拿掉這個位置的數量">${icon('x', 18)}</button></div>`).join('')}</div>`
        : '<div class="group"><div class="row muted">還沒盤點過。盤點時按「完成」就會記在這裡。</div></div>'
    }
    </div><div class="detail-side">
    <p class="section-title">叫貨提醒</p>
    <div class="group"><div class="row"><span class="grow"><span class="title">剩幾個就要叫貨</span><br><span class="meta">${needsOrder(it) ? `<span class="warn-line">${icon('warning', 16)}<span>${esc(orderWhy(it))}，該叫貨了</span></span>` : it.safety == null ? '例如設 3：剩 3 個以下就列進「叫貨」（安全庫存）' : `剩 ${it.safety} 個以下，就列進「叫貨」`}</span></span><span class="stepper edit-only"><button data-safety="-1" aria-label="減一">${icon('minus', 20)}</button><input id="safety" inputmode="numeric" value="${it.safety ?? ''}" placeholder="—" aria-label="剩幾個就要叫貨"><button data-safety="1" aria-label="加一">${icon('plus', 20)}</button></span><span class="qty view-only"><b>${it.safety ?? '—'}</b></span></div></div>
    ${
      decoded.length
        ? `<p class="section-title">型號解讀</p><div class="group">${decoded.map((x) => `<div class="row decode"><span class="grow"><span class="title">${esc(x.title)}</span>${x.facts.map((t) => `<br><span class="meta">・${esc(t)}</span>`).join('')}</span></div>`).join('')}</div>`
        : ''
    }
    <p class="section-title">替代品（可以互換）</p>
    <div class="group">
      ${eq.rule.map((e) => equivRow(e)).join('')}
      ${eq.linked.map((x) => equivRow({ brand: x.brand, model: x.model || itemTitle(x), item: x, linked: true })).join('')}
      <button class="row edit-only" data-action="equiv-add"><span class="add-dot" aria-hidden="true">${icon('plus', 18)}</span><span class="grow"><span class="title" style="color:var(--tint)">加一個可以互換的品項</span><br><span class="meta">例如客人常問的別牌同規格</span></span></button>
    </div>
    ${decoded.find((x) => x.note)?.note ? `<p class="footnote">${esc(decoded.find((x) => x.note).note)}</p>` : ''}
    ${links.length ? `<p class="section-title">查原廠資料</p><div class="group">${links.map((l) => `<a class="row" href="${esc(l.url)}" target="_blank" rel="noreferrer"><span class="grow">${esc(l.title)}</span>${chev}</a>`).join('')}</div>` : ''}
    ${
      it.moves?.length
        ? `<p class="section-title">進出紀錄</p><div class="group">${[...it.moves]
            .reverse()
            .slice(0, 20)
            .map((m) => `<div class="row"><span class="grow">${m.kind === 'in' ? '進貨' : m.kind === 'out' ? '賣出' : m.from === 'undo' ? '帳面數改回原本的' : m.why ? '複盤後調整帳面' : '設定帳面數'}<br><span class="meta">${fmtTime(m.at)}${m.why ? `・${esc(m.why)}` : ''}</span></span>${m.kind === 'set' && blindMe() ? '' : `<span class="qty"><b>${m.kind === 'in' ? '+' : m.kind === 'out' ? '−' : '＝'}${m.qty == null ? '—' : esc(m.qty)}</b></span>`}</div>`)
            .join('')}</div>`
        : ''
    }
    </div></div>
    <div class="row-actions edit-only" style="margin-top:22px"><button class="btn small secondary" data-action="item-merge">合併到另一個品項</button><button class="btn small danger" data-action="item-delete">刪除品項</button></div>
    <p class="footnote">以前的寫法（AI 認過的名稱）：${esc((it.aliases || []).length)} 種，以後辨識到都會算進這一項。</p>
  </main>`
}
/** 品項頁的「複盤」：每一張單的狀態、誰數的、原因；照權限顯示能按的按鈕 */
function recountSection(it) {
  const list = recountsOf(it).slice(0, 4)
  if (!list.length) return ''
  const blind = blindMe()
  const who = (c) => [c.by || '沒有記錄盤點人', c.at ? fmtTime(c.at) : ''].filter(Boolean).join('・')
  const rows = list
    .map(([sid, rc]) => {
      const counts = [`第一次：${esc(who({ ...rc.first, at: rc.at }))}${blind ? '' : `，${rc.first.count} 個`}`, ...rc.counts.map((c, i) => `複盤${rc.counts.length > 1 ? ` ${i + 1}` : ''}：${esc(who(c))}${blind ? '' : `，${c.count} 個`}`)]
      const reason = rc.reason ? `原因：${esc(rc.reason)}${rc.note ? `（${esc(rc.note)}）` : ''}` : ''
      const btns = []
      if (rc.status === 'pending') btns.push(`<button class="btn small edit-only" data-action="recount" data-sid="${esc(sid)}">我來複盤</button>`)
      if (rc.status === 'confirmed') {
        btns.push(`<button class="btn small secondary edit-only" data-action="recount-reason" data-sid="${esc(sid)}">${rc.reason ? '改原因' : '選原因'}</button>`)
        if (canManage()) btns.push(`<button class="btn small edit-only" data-action="recount-adjust" data-sid="${esc(sid)}">調整帳面</button><button class="btn small plain edit-only" data-action="recount-keep" data-sid="${esc(sid)}">不調整</button>`)
      }
      const tag = rcOpen(rc) ? 'low' : rc.status === 'fixed' ? 'ok' : 'edit'
      return `<div class="row" style="flex-wrap:wrap"><span class="grow"><span class="title">${placeHtml(rc.place)}</span><span class="badge ${tag}">${RC_TEXT[rc.status]}</span><br><span class="meta">${counts.join('<br>')}${reason ? `<br>${reason}` : ''}${rc.status === 'pending' && rc.counts.length ? '<br>兩次數字不一樣：請再找一個人數一次' : ''}${rc.status === 'confirmed' && !canManage() ? '<br>等擁有者或管理員決定要不要調整帳面' : ''}</span></span>${btns.length ? `<span class="row-actions" style="width:100%;margin-top:8px">${btns.join('')}</span>` : ''}</div>`
    })
    .join('')
  return `<p class="section-title">複盤</p><div class="group">${rows}</div>`
}

function equivRow(e) {
  const x = e.item
  return x
    ? `<button class="row" data-item-open="${esc(x.id)}"><span class="grow"><span class="title">${esc([e.brand, e.model].filter(Boolean).join(' '))}</span><br><span class="meta in-stock">店裡有 ${onHand(x)}・${esc(x.no)}${e.linked ? '・自己設定的' : ''}</span></span>${chev}</button>`
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
  // 店裡沒有一模一樣的：找型號相近的（例如查 DML 083S，店裡有 DML 083）
  const qa = alnum(read?.model || q)
  const near =
    !matches.length && qa.length >= 4
      ? items
          .filter((it) => {
            const m = alnum(it.model)
            return m.length >= 4 && (qa.startsWith(m) || m.startsWith(qa) || (m.length >= 5 && qa.slice(0, -1) === m.slice(0, -1)))
          })
          .slice(0, 10)
      : []
  // 型號相近、但看得出尺寸或規格不一樣（例如 DML 083＝3分、DML 084＝4分）：明講不能互換
  const specOf = (m) => decode(m).find((x) => x.model)?.spec || ''
  const sizeOf = (s) => (String(s).match(/\d+吋\d+分|\d+分/) || [''])[0]
  const qSpec = specOf(read?.model || q)
  const nearNote = (it) => {
    const s = specOf(it.model) || it.spec || ''
    if (!qSpec || !s || canon(s) === canon(qSpec)) return ''
    const a = sizeOf(qSpec)
    const b = sizeOf(s)
    return a && b && a !== b ? `尺寸不同（${a}／${b}），不能互換` : `規格不同（${qSpec}／${s}），不能互換`
  }
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
            <div class="read-head">${read.url ? zoomWrap(`<img src="${read.url}" alt="拍到的標籤">`, zoomAttrs(read.url, ['拍到的標籤', read.label, read.model].filter(Boolean).join('・')), 'lg') : ''}<div class="grow"><span class="stock-label">AI 讀到</span><b>${esc(read.label || '（看不出品名）')}</b><span class="meta">${esc([read.brand, read.model, read.spec].filter(Boolean).join('・') || '看不出型號')}</span>${read.code ? `<br><span class="meta">訂購碼 ${esc(read.code)}</span>` : ''}</div></div>
            ${read.text ? `<details class="trace"><summary>標籤上的字</summary>${esc(read.text)}</details>` : ''}
            <div class="row-actions" style="margin-top:12px">${exact ? `<button class="btn small" data-item-open="${esc(exact.id)}">打開 ${esc(exact.no)}（店裡有 ${onHand(exact)}）</button>` : '<button class="btn small edit-only" data-action="read-add">加入品項庫</button>'}<button class="btn small secondary" data-action="read-clear">清除</button></div>
          </section>`
        : ''
    }
    ${decoded.length ? `<p class="section-title">這是什麼</p><div class="group">${decoded.map((x) => `<div class="row decode"><span class="grow"><span class="title">${esc(x.title)}</span>${x.facts.map((t) => `<br><span class="meta">・${esc(t)}</span>`).join('')}</span></div>`).join('')}</div>` : ''}
    <p class="section-title">店裡有的（${matches.length}）</p>
    ${matches.length ? `<div class="group">${matches.map(itemRow).join('')}</div>` : `<div class="group"><div class="row muted">品項庫裡沒有${fields.model || fields.label ? `「${esc(read ? read.model || read.label : q)}」` : ''}。</div></div>`}
    ${near.length ? `<p class="section-title">型號相近的（${near.length}）<span class="badge new">新</span></p><div class="group">${near.map((it) => itemRow(it) + (nearNote(it) ? `<div class="row near-note"><span class="warn-line">${icon('warning', 16)}<span>${esc(nearNote(it))}</span></span></div>` : '')).join('')}</div><p class="footnote">型號只差一點（例如尾巴多一個字母），不一定能互換；請看規格再確認。</p>` : ''}
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
    <div class="nav nav-empty"><span></span></div>
    <h1 class="large-title">查型號</h1>
    <p class="subtitle">客人拿零件或型號來問：拍標籤或打型號，馬上看是什麼、店裡有沒有、可以用什麼替代。</p>
    ${
      hasAi()
        ? `<label class="hero-btn ${busy ? 'busy' : ''}" ${busy ? 'aria-disabled="true"' : ''}><span class="hero-icon" aria-hidden="true">${busy ? '<span class="spinner small"></span>' : icon('camera', 30)}</span><span class="grow"><b>${busy ? 'AI 讀標籤中…' : '拍型號標籤'}</b><span class="meta">${busy ? '大約 5～15 秒' : '外盒、貼紙、機器上的型號牌、零件上的刻字都可以'}</span></span><input type="file" accept="image/*" capture="environment" id="label-cam" class="sr-only" ${busy ? 'disabled' : ''}></label>`
        : `<button class="hero-btn off" data-action="lookup-no-ai" aria-disabled="true"><span class="hero-icon" aria-hidden="true">${icon('camera', 30)}</span><span class="grow"><b>拍型號標籤</b><span class="meta">${!canEdit() && aiSharedOn() ? '檢視者不能用公司共用的金鑰' : '要先設定 AI 才能用；打型號查詢不用'}</span></span></button>`
    }
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
    <p class="subtitle">每一格貨架一個代號（例如 ${locTag('A-01', 'sm')}）。印成標籤貼在貨架上：拍照時 AI 看到標籤，就會自動填位置；重盤同一格，數量自動更新。</p>
    ${
      locs.length
        ? `<div class="group">${locs
            .map((l, i) => {
              const its = here(l.code)
              const qty = its.reduce((n, it) => n + inLoc(it, l.code).reduce((m, [, st]) => m + st.count, 0), 0)
              return `<button class="row loc-row" data-loc-edit="${i}"><span class="loc-slot">${locTag(l.code, 'lg')}</span><span class="grow"><span class="title">${esc(l.name || '（沒有說明）')}</span><br><span class="meta">${its.length ? `${its.length} 種・${qty} 件` : '還沒盤點'}</span></span>${chev}</button>`
            })
            .join('')}</div>`
        : (() => {
            // 還沒建立儲位、但盤點時填過位置：寫出用過哪些，跟品項庫的「用到 N 個位置」說法一致
            const used = [...new Set(items.flatMap((it) => liveStock(it).map(([, st]) => st.place)).filter(Boolean))]
            return `<div class="group"><div class="row muted"><span>還沒建立儲位。${used.length ? `盤點時用過 ${used.length} 個位置：<span class="loc-list">${used.slice(0, 8).map((p) => locTag(p, 'sm')).join('')}</span>${used.length > 8 ? '…' : ''}<br>` : ''}按下面「新增儲位」，可以一次建立一整排。</span></div></div>`
          })()
    }
    <div class="row-actions" style="margin-top:14px"><button class="btn edit-only" style="flex:1" data-action="loc-add">${icon('plus', 20)}新增儲位</button><button class="btn secondary" style="flex:1" data-action="loc-print" ${locs.length ? '' : 'disabled'}>${icon('tag', 20)}列印標籤</button></div>
    <details class="steps"><summary>怎麼編號比較好？</summary>
      <ol><li>字母＝第幾排貨架（A、B、C…），數字＝第幾層（由上往下 01、02…）。例：B-03＝B 排第 3 層。</li><li>標籤印出來剪下，貼在每一層的正中間、正面朝外；拍照時把標籤一起拍進去。</li><li>同一格要一次拍完（可以拍好幾張）。重盤同一格，會以新的那次為準（這次沒拍到的，就從這一格拿掉）；不同格的數量會加起來變成「實盤」。</li></ol>
    </details>
  </main>`
}

// ───────────────────────── 點貨對單（4.7 測試版） ─────────────────────────
/**
 * 進貨（廠商送來）、出貨（送客戶前撿貨）時對單對貨：名稱對、數量對，不出錯。整個功能標「測試版」。
 * 一張點貨單（idb.picks）：{ id, kind:'in'|'out', ref 單號（選填）, createdAt, doneAt 第一次按「完成」的時間,
 *   editedAt／editedBy 最後一次「改一下」的時間和人, by／byId 誰點的,
 *   lines:[{ id, no 正航產品編號 | itemId 品項庫的品項（還沒接正航的）, name／code／unit 加進來時的名稱（產品表換了也看得懂）,
 *            qty 單子數量, got 實拿數量（空白＝跟單子一樣）, done 點好了, guess AI 只靠名稱配的、還沒確認,
 *            cands 對不到時有幾個像的, st 按完成時的狀態, read AI 讀到的字 }] }
 * 數量可以有小數（最多兩位：冷媒 kg、銅管 M）。不存客戶名稱、價格；點貨不改帳面（帳面以正航為準），只留紀錄。
 * 按「完成」後才同步給大家（鍵 pick:ID）。「改一下」是暫存的副本：按「完成修改」才寫回、才同步。
 */
const PICK_KIND = { in: '進貨', out: '出貨' }
const PICK_TITLE = { in: '進貨點貨', out: '出貨撿貨' }
const PICK_MAX_LINES = 200
const PICK_QTY_MAX = 99999
/** 文字長度上限（加進來、AI 讀到、別台同步來的都用同一套） */
const PICK_LEN = { ref: 40, name: 160, code: 60, spec: 60, unit: 12, by: 40, id: 60, readCode: 40, readName: 80 }
/** 交叉比對最多列幾項（再多的只寫數量） */
const PICK_CROSS_MAX = 300
/** 一張單子最多遮幾欄 */
const PICK_MASK_MAX = 6
/** 4.7 上線前一天：更新後從這一天之後的同步紀錄再拿一次（補舊版略過的點貨紀錄） */
const PICK_LAUNCH = Date.parse('2026-10-09T00:00:00+08:00')
/** 同步來的時間要在合理範圍（2020 年到現在＋1 天） */
const PICK_TIME_MIN = Date.parse('2020-01-01T00:00:00+08:00')
const PICK_STATE = { todo: '還沒點', ok: '對了', bad: '數量不對', none: '對不到產品', guess: 'AI 配的，請確認' }
const betaBadge = '<span class="badge beta">測試版</span>'
/** 交叉比對（測試版）：開關只在這台、只有擁有者／管理員 */
const crossOn = () => canManage() && ls.get(LS.pickCross) === '1'
/** 有按完成的點貨單改過、還沒檢查要不要上傳：同步時才讀點貨紀錄（不用每 40 秒全部讀一遍） */
let picksDirty = true

// —— 對照、比對、數量（純函式：node 驗證腳本直接從這個檔案取出來測） ——
/** AI 讀到的字、別台傳來的字一律當資料：拿掉控制字元、空白合併、限制長度 */
function clip(s, n) {
  return String(s ?? '')
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, n)
}
const round2 = (n) => Math.round(n * 100) / 100
/** 合理的數量：0～99999，最多兩位小數 */
const qtyOk = (n) => typeof n === 'number' && Number.isFinite(n) && n >= 0 && n <= PICK_QTY_MAX && Math.abs(round2(n) - n) < 1e-9
/** 打的數量（可以有小數、全形也認得）：「2.5」→ 2.5；負數、看不懂的 → null；超過兩位小數四捨五入 */
function parseQty(v) {
  const s = String(v ?? '')
    .normalize('NFKC')
    .replace(/[,\s]/g, '')
  if (!/^(\d+(\.\d*)?|\.\d+)$/.test(s)) return null
  const n = round2(Number(s))
  return Number.isFinite(n) && n <= PICK_QTY_MAX ? n : null
}
const fmtQty = (n) => (n == null || n === '' ? '' : String(round2(Number(n))))
/** 實拿：空白＝跟單子一樣（已點好的、狀態、結果頁都用這個，才不會一個說對、一個說錯） */
const gotOf = (l) => l.got ?? l.qty
function splitWords(text) {
  return String(text ?? '')
    .split(/[\s,，、;；]+/)
    .filter((w) => canon(w))
}
/**
 * 點貨用的對照表：正航產品＋品項庫裡還沒接正航的品項。
 * byExact：編號原樣 → 那一筆；byCanon：編號整理後（大小寫、橫線不計）→ 全部符合的（T-0404、T0404 兩種都留著，不會從搜尋消失）。
 * rows：搜尋用（每一筆的字先整理好 h）；已經接上正航的品項算在那個正航產品裡（照片、位置從品項來）。
 */
function pickIndex(products, items) {
  const byExact = new Map()
  const byCanon = new Map()
  const rows = []
  const linked = erpIndex(items || [])
  const taken = new Set()
  for (const p of products || []) {
    const k = canon(p.no)
    if (!k || byExact.has(p.no)) continue
    const it = linked.of(p.no)
    if (it) taken.add(it)
    const r = { no: p.no, p, it, h: normHay(`${p.no} ${p.name}`) }
    byExact.set(p.no, r)
    byCanon.set(k, [...(byCanon.get(k) || []), r])
    rows.push(r)
  }
  const itemByCanon = new Map()
  const itemById = new Map()
  for (const it of items || []) {
    itemById.set(it.id, it)
    if (taken.has(it)) continue
    const r = { it, h: normHay([it.no, it.label, it.brand, it.model, it.spec].join(' ')) }
    const k = canon(it.no)
    if (k) itemByCanon.set(k, [...(itemByCanon.get(k) || []), r])
    rows.push(r)
  }
  return { byExact, byCanon, itemByCanon, itemById, rows, linked }
}
const rowOfNo = (ix, no) => ix.byExact.get(no) || ix.byCanon.get(canon(no))?.[0] || null
/** 對照到的那一筆，記在點貨單上的樣子（名稱、編號、單位也記一份：之後產品表換了也看得懂） */
function pickTarget(r) {
  return r.no
    ? { no: r.no, name: clip(r.p.name || r.no, PICK_LEN.name), code: clip(r.no, PICK_LEN.code), unit: clip(r.p.unit, PICK_LEN.unit) }
    : { itemId: r.it.id, name: clip(itemTitle(r.it), PICK_LEN.name), code: clip(r.it.no, PICK_LEN.code), unit: clip(r.it.unit, PICK_LEN.unit) }
}
/**
 * 每一個詞都要「完整」出現在 h 裡（h 已經 normHay）：DML 083 不會對到 DML 083S、KP 1 不會對到 KP15、2.5L 不會對到 25L。
 * 相鄰的詞先試接起來（「DML 083S」＝「DML083S」），接不起來再一個一個比。ws：先拆好的詞（或一段字）
 */
function wordsIn(h, ws) {
  if (!Array.isArray(ws)) ws = splitWords(ws)
  if (!ws.length) return false
  for (let i = 0; i < ws.length; ) {
    let j = Math.min(ws.length, i + 3)
    for (; j > i; j--) if (tokenRe(ws.slice(i, j).join(' '))?.test(h)) break
    if (j === i) return false
    i = j
  }
  return true
}
/** 打到一半：前面的詞完整出現、最後一個詞「開頭一樣」就算。只用在「開頭一樣的」那一區，讓人自己選，不拿來自動配對 */
function prefixIn(h, ws) {
  if (!Array.isArray(ws)) ws = splitWords(ws)
  if (!ws.length) return false
  if (ws.length > 1 && !wordsIn(h, ws.slice(0, -1))) return false
  return !!tokenRe(ws[ws.length - 1], false)?.test(h)
}
/** 名稱有沒有關係：有一個詞完整出現，或中文名稱有兩個字連在一起出現（乾燥器／乾燥過濾器）；完全不相干＝編號剛好撞號 */
function related(h, text) {
  if (splitWords(text).some((w) => canon(w).length >= 2 && tokenRe(w)?.test(h))) return true
  const han = normHay(text).match(/[一-鿿]{2,}/g) || []
  return han.some((s) => [...s].some((_, i) => i < s.length - 1 && h.includes(s.slice(i, i + 2))))
}
/**
 * 搜尋產品（打產品編號、品名、型號）：
 * 1. 產品編號（料號）一模一樣的排第一（canon：大小寫、橫線、空白不計）
 * 2. 品名、型號每個詞都完整出現的（hasToken 的規則，不用會配錯的子字串比對）
 * 3. 還不夠：編號或型號「開頭一樣的」另外列（打到一半時用；不一定是同一個）
 */
function pickSearch(ix, q, limit = 8) {
  const raw = String(q ?? '').trim()
  const out = { hits: [], near: [], more: 0 }
  const cq = canon(raw)
  if (!cq) return out
  const ws = splitWords(raw)
  const seen = new Set()
  const add = (list, r) => {
    if (!r) return
    const k = r.no ? `n:${r.no}` : `i:${r.it.id}`
    if (seen.has(k)) return
    seen.add(k)
    list.push(r)
  }
  for (const r of ix.byCanon.get(cq) || []) add(out.hits, r)
  for (const r of ix.itemByCanon.get(cq) || []) add(out.hits, r)
  for (const r of ix.rows) if (wordsIn(r.h, ws)) add(out.hits, r)
  out.more = Math.max(0, out.hits.length - limit)
  out.hits = out.hits.slice(0, limit)
  if (out.hits.length < limit)
    for (const r of ix.rows) {
      if (out.near.length >= limit) break
      if (prefixIn(r.h, ws)) add(out.near, r)
    }
  return out
}
/**
 * AI 讀到的一列 → 哪一個產品。回傳 { r 配到的, sure 確定, cands 有幾個像的 }：
 * - 編號一樣、只有一個、名稱也有關（或單子上沒寫名稱）＝確定
 * - 編號一樣但名稱完全不相干（廠商編號剛好撞號）、或只靠名稱配到的＝AI 猜的，要人按「是這個」才算
 * - 名稱好幾個都像、或都不像＝對不到（cands 寫有幾個像的，讓人選）
 */
function matchRow(ix, row) {
  const text = [row?.name, row?.spec].filter(Boolean).join(' ')
  const c = canon(row?.code)
  if (c) {
    const hits = [...(ix.byCanon.get(c) || []), ...(ix.itemByCanon.get(c) || [])]
    if (hits.length) {
      const r = hits.find((x) => x.no === row.code) || hits[0]
      return { r, sure: hits.length === 1 && (!canon(text) || related(r.h, text)), cands: hits.length }
    }
  }
  if (canon(text).length < 2) return { r: null, sure: false, cands: 0 }
  const ws = splitWords(text)
  const hits = ix.rows.filter((x) => wordsIn(x.h, ws))
  return hits.length === 1 ? { r: hits[0], sure: false, cands: 1 } : { r: null, sure: false, cands: hits.length }
}
/** AI 回的 JSON → 品項列：字串一律當資料（限制長度）；數量要大於 0、最多兩位小數，不是就留空讓人補 */
function cleanNoteRows(parsed) {
  const rows = Array.isArray(parsed?.rows) ? parsed.rows : []
  return rows
    .slice(0, PICK_MAX_LINES)
    .filter((r) => r && typeof r === 'object')
    .map((r) => {
      const n = typeof r.qty === 'number' ? r.qty : parseQty(r.qty)
      return { code: clip(r.code, PICK_LEN.readCode), name: clip(r.name, PICK_LEN.readName), spec: clip(r.spec, PICK_LEN.spec), qty: n > 0 && qtyOk(n) ? n : null }
    })
    .filter((r) => r.code || r.name)
}
/**
 * 別台同步來的點貨單也當資料看（白名單）：不是物件就不要；id 一律用同步鍵上的 id；
 * 種類只認 in／out；數量 0～99999、最多兩位小數；時間在 2020 年到現在＋1 天；字串限制長度；每一列的 id 不重複。
 */
function normPick(d, id) {
  if (!d || typeof d !== 'object' || Array.isArray(d)) return null
  const key = String(id ?? d.id ?? '').slice(0, PICK_LEN.id)
  if (!key) return null
  const time = (v) => {
    const n = Number(v)
    return Number.isFinite(n) && n >= PICK_TIME_MIN && n <= Date.now() + 86400000 ? n : undefined
  }
  const qty = (v) => (qtyOk(v) ? round2(v) : null)
  const opt = (v, n) => (v == null || v === '' ? undefined : String(v).slice(0, n))
  const ids = new Set()
  return {
    id: key,
    kind: d.kind === 'in' ? 'in' : 'out',
    ref: clip(d.ref, PICK_LEN.ref),
    createdAt: time(d.createdAt) || time(d.doneAt) || 0,
    doneAt: time(d.doneAt),
    editedAt: time(d.editedAt),
    editedBy: clip(d.editedBy, PICK_LEN.by) || undefined,
    by: clip(d.by, PICK_LEN.by),
    byId: opt(d.byId, PICK_LEN.id),
    lines: (Array.isArray(d.lines) ? d.lines : [])
      .filter((l) => l && typeof l === 'object' && !Array.isArray(l))
      .slice(0, PICK_MAX_LINES)
      .map((l, i) => {
        let lid = opt(l.id, PICK_LEN.id) || `l${i}`
        if (ids.has(lid)) lid = `${lid}-${i}`
        ids.add(lid)
        return {
          id: lid,
          no: opt(l.no, PICK_LEN.code),
          itemId: opt(l.itemId, PICK_LEN.id),
          name: clip(l.name, PICK_LEN.name),
          code: clip(l.code, PICK_LEN.code),
          unit: clip(l.unit, PICK_LEN.unit),
          qty: qty(l.qty),
          got: qty(l.got),
          done: !!l.done,
          guess: l.guess ? true : undefined,
          cands: Number.isInteger(l.cands) && l.cands >= 0 && l.cands < 100000 ? l.cands : undefined,
          st: ['ok', 'bad', 'todo', 'none', 'guess'].includes(l.st) ? l.st : undefined,
          read: l.read && typeof l.read === 'object' ? { code: clip(l.read.code, PICK_LEN.readCode), name: clip(l.read.name, PICK_LEN.readName), spec: clip(l.read.spec, PICK_LEN.spec) } : undefined,
        }
      }),
  }
}
/** 一項的狀態：none 對不到產品、guess AI 配的還沒確認（都是琥珀）、todo 還沒點（灰）、ok 對了（綠）、bad 數量不對（紅） */
function lineState(l) {
  if (!l.no && !l.itemId) return 'none'
  if (l.guess) return 'guess'
  if (!l.done) return 'todo'
  const g = gotOf(l)
  return l.qty != null && g != null && Math.abs(g - l.qty) < 0.001 ? 'ok' : 'bad'
}
/** 點好了的那一項實拿幾個；沒點的、還沒確認產品的不算（不知道到底拿了什麼） */
function lineGot(l) {
  return l.done && !l.guess && (l.no || l.itemId) ? Number(gotOf(l)) || 0 : 0
}
/** 統計：none＝對不到（含 AI 配的還沒確認，guess 另外也算一份）；done＝已點（對的＋數量不對的） */
function pickSummary(p) {
  const s = { total: 0, done: 0, ok: 0, bad: 0, todo: 0, none: 0, guess: 0 }
  for (const l of p?.lines || []) {
    s.total++
    const st = lineState(l)
    if (st === 'guess') s.guess++
    s[st === 'guess' ? 'none' : st]++
  }
  s.done = s.ok + s.bad
  return s
}
/**
 * 兩台同時改同一張（「改一下」時另一台也改了）：用每一列的 id 合併。
 * base＝開始改的時候；mine＝這台改好的；theirs＝雲端（別台）的最新版。
 * 這台改過的列用這台的；這台沒動、別台改過或新加的用別台的；這台刪掉的不加回來；別台刪掉、這台沒動的就刪。
 */
function mergePick(base, mine, theirs) {
  const sig = (l) => JSON.stringify([l.no, l.itemId, l.qty, l.got, l.done, !!l.guess, l.name])
  const baseBy = new Map((base?.lines || []).map((l) => [l.id, l]))
  const theirsBy = new Map((theirs?.lines || []).map((l) => [l.id, l]))
  const lines = []
  const seen = new Set()
  for (const l of mine.lines || []) {
    seen.add(l.id)
    const b = baseBy.get(l.id)
    const t = theirsBy.get(l.id)
    const untouched = b && sig(b) === sig(l)
    if (untouched && t) lines.push(t)
    else if (untouched && !t) continue
    else lines.push(l)
  }
  for (const t of theirs?.lines || []) if (!seen.has(t.id) && !baseBy.has(t.id)) lines.push(t)
  return { ...theirs, ...mine, ref: (mine.ref || '') !== (base?.ref || '') ? mine.ref : theirs?.ref || '', lines }
}
/**
 * 交叉比對（測試版）：上次匯入正航（基準 { qty, at }）之後、這張報表之前有點過貨的產品，
 * 點貨的變化（進貨實拿總和 − 出貨實拿總和）≠ 正航的變化（這次數量 − 基準數量）→ 列進「要查一下」（容許 0.001 的誤差）。
 * 不說誰錯：可能單子打錯、拿錯貨，或有進出貨沒用 App 點。
 * oldProducts：上次的產品表（有 at 的才有基準）；newProducts：這次報表讀到的；picks：點貨紀錄（按過完成的才算）；now：報表的時間。
 */
function crossCheck(oldProducts, newProducts, picks, now) {
  const oldBy = new Map((oldProducts || []).map((p) => [p.no, p]))
  const first = !(oldProducts || []).some((p) => p.at)
  // 先建好「產品 → 哪幾張點貨單、各拿幾個」，不要每一種產品都把全部點貨單掃一遍
  const byNo = new Map()
  for (const pk of picks || []) {
    if (!pk?.doneAt || pk.doneAt > now) continue
    for (const l of pk.lines || []) {
      if (!l.no || l.guess) continue
      let m = byNo.get(l.no)
      if (!m) byNo.set(l.no, (m = new Map()))
      m.set(pk, (m.get(pk) || 0) + lineGot(l))
    }
  }
  const items = []
  const inNew = new Set()
  let checked = 0
  for (const np of newProducts || []) {
    inNew.add(np.no)
    const op = oldBy.get(np.no)
    const m = byNo.get(np.no)
    if (!op?.at || !m) continue
    const rel = [...m].filter(([pk]) => pk.doneAt > op.at).map(([pk, got]) => ({ id: pk.id, at: pk.doneAt, ref: pk.ref || '', kind: pk.kind, got: round2(got) }))
    if (!rel.length) continue
    checked++
    const pick = round2(rel.reduce((n, x) => n + (x.kind === 'in' ? x.got : -x.got), 0))
    const erp = round2((Number(np.qty) || 0) - (Number(op.qty) || 0))
    if (Math.abs(pick - erp) > 0.001) items.push({ no: np.no, name: np.name || op.name || '', pick, erp, picks: rel.sort((a, b) => a.at - b.at) })
  }
  // 有點過貨、這次的正航表裡卻沒有（報表只匯了一部分？）：沒辦法比，只算數量
  let missing = 0
  for (const [no, m] of byNo) {
    const op = inNew.has(no) ? null : oldBy.get(no)
    if (op?.at && [...m.keys()].some((pk) => pk.doneAt > op.at)) missing++
  }
  return { first, checked, missing, total: items.length, items: items.slice(0, PICK_CROSS_MAX) }
}

// —— 存檔 ——
let picksCache = null
async function picksAll(force = false) {
  if (!picksCache || force) picksCache = await idb.picks.all().catch(() => [])
  return picksCache
}
/**
 * 存點貨單：
 * - 「改一下」的暫存副本（_copy）：不存，按「完成修改」才寫回
 * - 已經刪掉的（_gone，例如刪掉時 AI 還在讀）：不再存回去
 * - 還沒按完成的：只存這台（不同步）；按過完成的：一般存檔，稍後同步給大家
 */
async function savePick(p) {
  if (p._gone || p._copy) return
  try {
    if (p.doneAt) {
      await idb.picks.put(p)
      picksDirty = true
    } else {
      p.updatedAt = Date.now()
      await idb.picks.putRaw(p)
    }
  } catch (e) {
    saveFailed(e)
    throw e
  }
  picksCache = null
}
/** 正在點的那一張：連續按 ＋／－ 時一次只存一個，存的時候又改了，存完再存一次（不會漏、也不會塞車） */
let pickSaving = null
let pickDirty = null
function pickSave(p = state.pick) {
  if (!p) return Promise.resolve()
  if (pickSaving) {
    pickDirty = p
    return pickSaving
  }
  pickSaving = savePick(p)
    .catch(() => {})
    .finally(() => {
      pickSaving = null
      const again = pickDirty
      pickDirty = null
      if (again) pickSave(again)
    })
  return pickSaving
}
const clonePick = (p) => JSON.parse(JSON.stringify(p))
const pickLine = (id) => state.pick?.lines.find((l) => l.id === id)
/** 現在是不是在「改一下」（暫存的副本） */
const editingCopy = () => !!(state.pick?._copy && state.pickEdit?.id === state.pick.id)
const whoAmI = () => currentCounter()?.name || ls.get(LS.memberName) || ''
function pickEditDirty() {
  if (!editingCopy()) return false
  const sig = (p) => JSON.stringify([p.ref || '', p.lines.map((l) => [l.id, l.no, l.itemId, l.qty, l.got, l.done, !!l.guess])])
  return sig(state.pick) !== sig(state.pickEdit.base)
}
/** 離開點貨單那一頁：「改一下」有改 → 先問要不要放棄；什麼都沒加的草稿不留。回傳 false＝不要離開 */
async function leavePickEdit() {
  const p = state.pick
  if (!p) return true
  if (editingCopy()) {
    if (pickEditDirty() && !confirm('放棄這次的修改？\n\n按「確定」放棄，紀錄維持原本的樣子；\n按「取消」回去繼續改，改好按「完成修改」。')) return false
    state.pickEdit = null
    state.pick = null
    return true
  }
  if (!p.doneAt && !p.lines.length && !p.ref) {
    p._gone = true
    await idb.picks.del(p.id).catch(() => {})
    picksCache = null
  }
  return true
}
/** 對照表：正航產品表換了、品項庫改過（改名、接上正航）才重建（4000 多種，每次重畫都建會慢） */
let pickIx = { erp: undefined, ver: -1, ix: null }
async function pickCtx() {
  const erp = await erpGet()
  const items = await itemsAll()
  if (pickIx.erp !== erp || pickIx.ver !== itemsVer) pickIx = { erp, ver: itemsVer, ix: pickIndex(erp?.products || [], items) }
  return pickIx.ix
}
const pickIxNow = () => pickIx.ix || pickIndex([], [])

// —— 畫面小零件 ——
/** 一項現在的名稱、編號、照片、位置：產品表或品項庫改了就用新的；都找不到就用加進來時記的 */
function lineInfo(ix, l) {
  const r = l.no ? rowOfNo(ix, l.no) : null
  const it = r?.it || (l.no ? ix.linked.of(l.no) : null) || (l.itemId ? ix.itemById.get(l.itemId) : null) || null
  const name = r?.p.name || (it ? itemTitle(it) : '') || l.name || l.read?.name || '（沒有名稱）'
  return { it, name, code: r?.no || it?.no || l.code || '', unit: r?.p.unit || it?.unit || l.unit || '' }
}
const readText = (l) => (l.read ? [l.read.code, l.read.name, l.read.spec].filter(Boolean).join(' ') : '')
const signed = (n) => {
  const v = round2(Number(n) || 0)
  return v > 0 ? `+${v}` : v < 0 ? `−${-v}` : '0'
}
function diffText(l, unit) {
  const d = round2((Number(gotOf(l)) || 0) - (Number(l.qty) || 0))
  const u = unit || '個'
  return d < 0 ? `少 ${-d} ${u}` : d > 0 ? `多 ${d} ${u}` : ''
}
function stateChip(l, unit) {
  const st = lineState(l)
  const text = st === 'bad' ? (l.qty == null ? '單子數量沒填' : diffText(l, unit) || PICK_STATE.bad) : PICK_STATE[st]
  return `<span class="pk-chip pk-state ${st}">${st === 'ok' ? icon('check', 16) : st === 'todo' ? '' : icon('warning', 16)}${esc(text)}</span>`
}
/** 數字框依位數縮字（4 位數以上不會被切掉） */
const lenAttr = (v) => ` data-len="${Math.min(9, String(v ?? '').length)}"`
const pickThumb = (info, size, inRow = false) => (info.it ? itemThumb(info.it, size, inRow) : `<span class="thumb ph" style="width:${size}px;height:${size}px" aria-hidden="true">${esc([...String(info.name || '?')][0] || '?')}</span>`)
/** 放在哪裡：只畫儲位黃標籤，不寫數量（點貨不需要帳面、實盤數，盲盤也一樣） */
function pickPlaces(info) {
  const pairs = info.it ? liveStock(info.it).map(([, s]) => [s.place, null]).filter(([pl]) => pl) : []
  return pairs.length ? placesHtml(pairs, 4) : '<span class="loc-list"><span class="loc-tag none sm">還沒記位置</span></span>'
}
function lineCard(ix, l, i) {
  const info = lineInfo(ix, l)
  const st = lineState(l)
  const id = esc(l.id)
  const rt = readText(l)
  const amber = st === 'none' || st === 'guess'
  const meta = [info.code, info.unit && `單位：${info.unit}`].filter(Boolean).join('・')
  // 對不到的：上面直接寫單子上的字（不重複寫兩次）；AI 配的：寫配到的產品，下面寫單子上的字讓人對
  const head =
    st === 'none'
      ? `<span class="pk-cap">單子上寫</span><span class="pk-name">${esc(rt || l.name || '（沒有名稱）')}</span>`
      : `<span class="pk-name">${esc(info.name)}</span>${meta ? `<span class="meta">${esc(meta)}</span>` : ''}${pickPlaces(info)}`
  const box =
    st === 'none'
      ? `<div class="pk-none"><p>${l.cands > 1 ? `有 ${l.cands} 個像的，請選一個。` : '正航產品表、品項庫裡找不到一樣的：請選是哪一個產品，或刪掉這項。'}</p><div class="row-actions"><button class="btn small" data-action="pk-assign" data-id="${id}">${icon('search', 18)}選產品</button><button class="btn small plain" data-action="pk-remove" data-id="${id}">刪掉這項</button></div></div>`
      : st === 'guess'
        ? `<div class="pk-none"><p>AI 配的，請確認${rt ? `：單子上寫「${esc(rt)}」` : ''}。</p><div class="row-actions"><button class="btn small" data-action="pk-confirm" data-id="${id}">${icon('check', 18)}是這個</button><button class="btn small plain" data-action="pk-assign" data-id="${id}">換一個</button></div></div>`
        : ''
  const need = fmtQty(l.qty) || '？'
  return `<article class="pk-line st-${st}" data-line="${id}" aria-label="第 ${i + 1} 項">
    <div class="pk-top">
      ${pickThumb(st === 'none' ? { it: null, name: rt || l.name } : info, 56)}
      <div class="grow">${stateChip(l, info.unit)}${head}</div>
      <button class="icon-btn small pk-more" data-action="pk-more" data-id="${id}" aria-label="這一項的其他動作：改單子數量、換產品、改回還沒點、刪掉" title="改數量、換產品、刪掉">${icon('more', 22)}</button>
    </div>
    ${box}
    <div class="pk-bot">
      <div class="pk-need"><span class="pk-cap">單子</span><button class="pk-need-num" data-action="pk-qty" data-id="${id}"${lenAttr(need)} aria-label="單子上寫 ${esc(need)}，點一下可以改">${esc(need)}</button></div>
      <div class="pk-got"><span class="pk-cap">實拿</span><span class="stepper pk-step"><button data-action="pk-step" data-id="${id}" data-d="-1" aria-label="少一個">${icon('minus', 22)}</button><input data-pk-got="${id}" inputmode="decimal" autocomplete="off" value="${esc(fmtQty(l.got))}" placeholder="${esc(fmtQty(l.qty))}"${lenAttr(fmtQty(l.got) || fmtQty(l.qty))} aria-label="實拿幾個（空白＝跟單子一樣）"><button data-action="pk-step" data-id="${id}" data-d="1" aria-label="多一個">${icon('plus', 22)}</button></span></div>
      <button class="pk-done${l.done ? ' on' : ''}${amber ? ' off' : ''}" data-action="pk-done" data-id="${id}" aria-pressed="${!!l.done}"${amber ? ' aria-disabled="true"' : ''}>${icon('check', 20)}<span>點好了</span></button>
    </div>
  </article>`
}
function pickProgress(p) {
  const s = pickSummary(p)
  const pct = s.total ? Math.round((s.done / s.total) * 100) : 0
  const allOk = s.total > 0 && s.ok === s.total
  return `<div class="pk-prog-text"><b>已點 ${s.done}／${s.total} 項</b>${s.bad ? `<span class="pk-chip bad">${s.bad} 項不對</span>` : ''}${s.none ? `<span class="pk-chip none">${s.none} 項要確認產品</span>` : ''}${allOk ? `<span class="pk-chip ok">${icon('check', 14)}全部對了</span>` : ''}</div><div class="pk-bar${allOk ? ' ok' : ''}" role="progressbar" aria-label="點貨進度" aria-valuemin="0" aria-valuemax="${s.total}" aria-valuenow="${s.done}"><span style="width:${pct}%"></span></div>`
}
/** 按 ＋／－、點好了、打數字：只更新那一張卡片和上面的進度（不整頁重畫，打字不會被打斷） */
function paintLine(l) {
  const card = document.querySelector(`.pk-line[data-line="${CSS.escape(l.id)}"]`)
  if (!card || !pickIx.ix) return render()
  const st = lineState(l)
  card.className = `pk-line st-${st}`
  const chip = card.querySelector('.pk-state')
  if (chip) chip.outerHTML = stateChip(l, lineInfo(pickIx.ix, l).unit)
  const input = card.querySelector('[data-pk-got]')
  if (input) {
    if (document.activeElement !== input) input.value = fmtQty(l.got)
    input.dataset.len = String(Math.min(9, (input.value || input.placeholder).length))
  }
  const btn = card.querySelector('.pk-done')
  if (btn) {
    btn.classList.toggle('on', !!l.done)
    btn.setAttribute('aria-pressed', String(!!l.done))
  }
  const prog = document.getElementById('pk-progress')
  if (prog) prog.innerHTML = pickProgress(state.pick)
}
/** 閃一下那一張卡片；scroll＝false：不捲過去（加品項時留在搜尋框） */
function flashLine(id, { scroll = true } = {}) {
  const card = document.querySelector(`.pk-line[data-line="${CSS.escape(id)}"]`)
  if (!card) return
  if (scroll) card.scrollIntoView({ block: 'center', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' })
  document.querySelectorAll('.pk-line.flash').forEach((x) => x.classList.remove('flash'))
  void card.offsetWidth
  card.classList.add('flash')
}
/** 搜尋結果（加品項、選產品共用）；mode＝'add'：點了加進清單；'assign'：點了＝這一項是這個產品 */
function pickResultsHtml(ix, q, mode = 'add') {
  if (!canon(q)) return ''
  const res = pickSearch(ix, q, 8)
  const row = (r) => {
    const t = pickTarget(r)
    const info = { it: r.it || null, name: t.name }
    return `<button class="row pk-res" ${mode === 'add' ? 'data-action="pk-add"' : 'data-pp=""'} ${r.no ? `data-no="${esc(r.no)}"` : `data-iid="${esc(r.it.id)}"`}>${pickThumb(info, 40, true)}<span class="grow"><span class="title">${esc(t.name)}</span><br><span class="meta">${esc([t.code, t.unit].filter(Boolean).join('・') || '品項庫')}</span>${info.it ? pickPlaces(info) : ''}</span>${mode === 'add' ? `<span class="pk-add-ico" aria-hidden="true">${icon('plus', 20)}</span>` : chev}</button>`
  }
  return `${res.hits.length ? `<div class="group">${res.hits.map(row).join('')}</div>${res.more ? `<p class="footnote">還有 ${res.more} 個也符合：打完整一點會更準。</p>` : ''}` : ''}${
    res.near.length ? `<p class="section-title">開頭一樣的（不一定是同一個）</p><div class="group">${res.near.map(row).join('')}</div>` : ''
  }${!res.hits.length && !res.near.length ? `<div class="group"><div class="row muted">找不到「${esc(String(q).trim().slice(0, PICK_LEN.ref))}」。${ix.byExact.size ? '請確認編號、型號有沒有打錯。' : '還沒匯入正航產品表：只找得到品項庫裡的。'}</div></div>` : ''}`
}
function pickRow(p) {
  const s = pickSummary(p)
  const by = byName(p)
  const meta = [fmtTime(p.doneAt || p.createdAt), by && `${by} 點`, `${s.total} 項`, p.editedAt && '改過'].filter(Boolean).join('・')
  const tag = !p.doneAt
    ? `<span class="pk-chip todo">點到 ${s.done}／${s.total}</span>`
    : !s.total
      ? '<span class="pk-chip todo">沒有品項</span>'
      : s.bad
        ? `<span class="pk-chip bad">${s.bad} 項不對</span>`
        : s.none
          ? `<span class="pk-chip none">${s.none} 項對不到</span>`
          : s.todo
            ? `<span class="pk-chip todo">${s.todo} 項還沒點</span>`
            : `<span class="pk-chip ok">${icon('check', 14)}全部對</span>`
  const kind = p.kind === 'in' ? 'in' : 'out'
  return `<button class="row pk-row" data-action="pk-open" data-id="${esc(p.id)}"><span class="pk-kind-ico ${kind}" aria-hidden="true">${icon(kind, 22)}</span><span class="grow"><span class="title">${PICK_KIND[kind]}・${p.ref ? esc(p.ref) : '<span class="muted">沒填單號</span>'}</span><br><span class="meta">${esc(meta)}</span></span>${tag}${chev}</button>`
}
function crossCard(last) {
  const when = last ? `・${fmtTime(last.at)}` : ''
  const n = last?.total ?? last?.items?.length ?? 0
  const meta = !last ? '匯入正航產品表時會自動比對' : last.first ? `下次匯入才開始比對${when}` : n ? `上次匯入有 ${n} 項要查一下${when}` : `上次匯入沒有要查的${when}`
  return `<div class="group pk-cross-card"><button class="row" data-action="pk-cross-last"><span class="row-ico" aria-hidden="true">${icon('chart')}</span><span class="grow"><span class="title">交叉比對（測試版）</span><span class="badge new">新</span><br><span class="meta">${esc(meta)}</span></span>${n ? `<span class="pk-chip bad">${n} 項</span>` : ''}${chev}</button></div>`
}
/** 設定頁：交叉比對開關（只有擁有者、管理員看得到） */
function pickCrossSettings() {
  const last = readJson(LS.pickCrossLast, null)
  const n = last?.total ?? last?.items?.length ?? 0
  return `<p class="section-title">點貨（測試版） <span class="badge new">新</span></p>
    <div class="group">
      <label class="row"><span class="grow"><span class="title">匯入正航時，跟點貨紀錄交叉比對（測試版）</span><br><span class="meta">只比上次匯入後有點過貨的產品：點貨記的進出跟正航的進出不一樣，就列出來查一下。第一次只記基準，下次匯入才開始比。開關只在這台。</span></span><input type="checkbox" id="rule-pick-cross" ${ls.get(LS.pickCross) === '1' ? 'checked' : ''} style="width:22px;height:22px"></label>
      ${last ? `<button class="row" data-action="pk-cross-last"><span class="grow"><span class="title" style="color:var(--tint)">看上次的比對結果</span><br><span class="meta">${esc(fmtTime(last.at))}・${esc(last.first ? '下次匯入才開始比對' : `${n} 項要查一下`)}</span></span>${chev}</button>` : ''}
    </div>`
}

// —— 三個畫面：點貨（清單）、點貨單（對貨）、結果 ——
async function viewPick() {
  const all = await picksAll()
  // 什麼都沒加的草稿（以前留下的）：順手清掉，不顯示
  const empty = all.filter((p) => !p.doneAt && !p.lines.length && !p.ref)
  if (empty.length) {
    for (const p of empty) idb.picks.del(p.id).catch(() => {})
    picksCache = null
  }
  const drafts = all.filter((p) => !p.doneAt && !empty.includes(p))
  const done = all.filter((p) => p.doneAt).sort((a, b) => b.doneAt - a.doneAt)
  const shown = done.slice(0, 30 + (state.pickMore || 0))
  const open = all.length - empty.length ? '' : 'open'
  const howTo = canEdit()
    ? `<details class="steps pk-howto" ${open}><summary>怎麼用？</summary>
      <ol>
        <li><b>開單：</b>按「出貨」或「進貨」；有單號就填（可以不填，不要填客戶名稱）。</li>
        <li><b>加品項：</b>打產品編號、搜品名，或按「拍單子讓 AI 讀」：先框出品項、遮住單價和金額那幾欄，只會送框裡、沒遮住的部分。</li>
        <li><b>點貨：</b>一項一項拿。數量跟單子一樣，直接按「點好了」；不一樣，先用 ＋／－ 改成實拿的數量，再按「點好了」。</li>
        <li><b>看顏色：</b>綠色＝對了，紅色＝數量不對（寫出差幾個），琥珀色＝對不到產品或 AI 配的要確認，灰色＝還沒點。</li>
        <li><b>完成：</b>按「完成」看結果，可以截圖或複製文字傳給老闆。點貨不會改帳面，只留紀錄。</li>
      </ol>
    </details>`
    : `<details class="steps pk-howto" ${open}><summary>怎麼看紀錄？</summary>
      <ol>
        <li><b>看顏色：</b>綠色「全部對」、紅色「幾項不對」、琥珀色「幾項對不到」、灰色「還沒點」。</li>
        <li><b>點進去看結果：</b>哪幾項數量不對、差幾個，誰點的、幾點點的。</li>
        <li><b>要傳給老闆：</b>直接截圖，或在結果頁按「複製結果」。</li>
      </ol>
    </details>`
  return `
  <main class="app">
    <div class="nav nav-empty"><span></span></div>
    <h1 class="large-title">點貨 ${betaBadge}</h1>
    <p class="subtitle">進貨、出貨時對單對貨：名稱對、數量對。只留點貨紀錄，不會改帳面（帳面以正航為準）。</p>
    <div class="pk-home">
      <div class="pk-home-main">
        ${
          canEdit()
            ? `<div class="pk-start">
                <button class="pk-start-btn" data-action="pk-new" data-kind="out"><span class="pk-start-ico" aria-hidden="true">${icon('out', 26)}</span><b>出貨</b><span class="meta">送客戶前撿貨<br>對出貨單</span></button>
                <button class="pk-start-btn" data-action="pk-new" data-kind="in"><span class="pk-start-ico" aria-hidden="true">${icon('in', 26)}</span><b>進貨</b><span class="meta">廠商送來<br>對進貨單</span></button>
              </div>`
            : `<div class="hint-card"><b>你是檢視者（只能看）</b>：可以看大家的點貨紀錄和結果；不能新增點貨單。需要點貨，請找擁有者或管理員改成「編輯者」。</div>`
        }
        ${drafts.length ? `<p class="section-title">還沒點完的（${drafts.length}）</p><div class="group">${drafts.map(pickRow).join('')}</div>` : ''}
        ${crossOn() ? crossCard(readJson(LS.pickCrossLast, null)) : ''}
        ${howTo}
      </div>
      <div class="pk-home-side">
        <p class="section-title">最近的點貨紀錄${done.length ? `（${done.length}）` : ''}</p>
        ${
          done.length
            ? `<div class="group">${shown.map(pickRow).join('')}</div>${done.length > shown.length ? `<button class="btn secondary block" data-action="pk-more-list" style="margin-top:10px">再顯示 ${Math.min(30, done.length - shown.length)} 筆（還有 ${done.length - shown.length} 筆）</button>` : ''}`
            : `<div class="empty pk-empty"><span class="empty-ico" aria-hidden="true">${icon('list-check', 34)}</span><p><b>還沒有點貨紀錄</b><br>${canEdit() ? '按「出貨」或「進貨」開一張，照單子一項一項點。' : '別人點完、同步後，紀錄會出現在這裡。'}</p>${canEdit() ? '<button class="btn small" data-action="pk-new" data-kind="out">開一張出貨單試試</button>' : ''}</div>`
        }
      </div>
    </div>
  </main>
  ${tabBar('pick')}`
}
async function viewPickEdit() {
  const p = state.pick
  if (!p) return viewPick()
  if (!canEdit()) {
    state.view = 'pick-result'
    return viewPickResult()
  }
  const ix = await pickCtx()
  const copy = editingCopy()
  const busy = state.pickBusyId === p.id
  const by = byName(p)
  const ai = busy
    ? `<div class="pk-ai-busy" role="status"><span class="spinner small" aria-hidden="true"></span><span class="grow"><b>AI 讀單子中…</b><br><span class="meta">大約 5～20 秒；只送框裡、沒遮住的部分</span></span></div>`
    : hasAi()
      ? `<div class="pk-ai"><div class="row-actions"><label class="btn secondary pk-ai-btn">${icon('camera', 20)}拍單子讓 AI 讀<input type="file" accept="image/*" capture="environment" id="pk-cam" class="sr-only"></label><label class="btn secondary pk-ai-btn">${icon('image', 20)}從相簿選<input type="file" accept="image/*" id="pk-album" class="sr-only"></label></div><p class="pk-privacy">${icon('shield', 16)}<span>拍完先框出品項、遮住單價和金額那幾欄；只會送框裡、沒遮住的部分給 AI。</span></p></div>`
      : `<div class="pk-ai off"><div class="row-actions"><button class="btn secondary pk-ai-btn" data-action="pk-no-ai" aria-disabled="true">${icon('camera', 20)}拍單子讓 AI 讀</button></div><p class="pk-privacy">${icon('warning', 16)}<span>${syncReady() ? '要先設定 AI：請擁有者到「設定」把 AI 金鑰放到雲端，或在「設定」貼上自己的 Gemini API Key。' : '要先設定 AI：到「設定」貼上 Gemini API Key。'}打編號、搜品名不用 AI。</span></p></div>`
  const right = copy
    ? `<button class="btn small plain" data-action="pk-cancel-edit">放棄修改</button>`
    : p.doneAt
      ? ''
      : `<button class="btn small plain" data-action="pk-discard" ${busy ? 'disabled' : ''}>刪掉這張</button>`
  return `
  <main class="app">
    <div class="nav">${backBtn('pick', '點貨')}<span class="nav-right">${right}</span></div>
    <h1 class="large-title">${PICK_TITLE[p.kind] || '點貨'} ${copy ? '<span class="badge edit">改一下</span>' : ''}${betaBadge}</h1>
    ${copy ? '<p class="subtitle">改的是暫存的副本：按「完成修改」才會存進紀錄、同步給大家；不改了按「放棄修改」。</p>' : ''}
    <div class="pk-cols">
      <div class="pk-addcol">
        <div class="group pk-meta">
          <label class="row"><span class="pk-label">單號</span><input class="inline" id="pk-ref" value="${esc(p.ref || '')}" placeholder="選填，例如 S1131010-001" maxlength="${PICK_LEN.ref}" autocomplete="off" spellcheck="false" enterkeyhint="done"></label>
          ${by ? `<div class="row"><span class="pk-label">點貨人</span><span class="grow">${esc(by)}</span><span class="meta">${esc(fmtTime(p.createdAt))}</span></div>` : ''}
        </div>
        <p class="footnote pk-ref-tip">單號只填單據號碼，不要填客戶名稱。</p>
        <p class="section-title">加品項</p>
        ${
          ix.rows.length
            ? `<input class="field search" id="pk-q" type="search" placeholder="打產品編號，或品名、型號" autocomplete="off" enterkeyhint="search" spellcheck="false" value="${esc(state.pickQ || '')}"><div id="pk-results">${pickResultsHtml(ix, state.pickQ || '')}</div>`
            : `<div class="hint-card">還沒匯入正航產品表，品項庫也是空的：請擁有者或管理員到「品項 → 產品總表」匯入，才找得到產品。</div>`
        }
        ${ai}
      </div>
      <div class="pk-linecol">
        ${
          p.lines.length
            ? `<div class="pk-progress" id="pk-progress">${pickProgress(p)}</div><div class="pk-lines">${p.lines.map((l, i) => lineCard(ix, l, i)).join('')}</div>`
            : `<div class="empty pk-empty"><span class="empty-ico" aria-hidden="true">${icon('list-check', 34)}</span><p><b>還沒有品項</b><br>把單子上的品項加進來：打產品編號、搜品名，或拍單子讓 AI 讀。</p></div>`
        }
      </div>
    </div>
  </main>
  <div class="toolbar"><div class="inner"><button class="btn" data-action="pk-finish" ${p.lines.length && !busy ? '' : 'disabled'}>${copy ? '完成修改' : '完成'}</button></div></div>`
}
async function viewPickResult() {
  const p = state.pick
  if (!p) return viewPick()
  const ix = await pickCtx()
  const s = pickSummary(p)
  const rows = p.lines.map((l) => ({ l, info: lineInfo(ix, l), st: lineState(l) }))
  const order = { bad: 0, guess: 1, none: 2, todo: 3 }
  const probs = rows.filter((r) => r.st !== 'ok').sort((a, b) => order[a.st] - order[b.st])
  const oks = rows.filter((r) => r.st === 'ok')
  const allOk = s.total > 0 && s.ok === s.total
  const by = byName(p)
  const kind = p.kind === 'in' ? 'in' : 'out'
  const head = [p.ref ? `單號 ${p.ref}` : '沒填單號', p.doneAt ? `${fmtTime(p.doneAt)} 完成` : `${fmtTime(p.createdAt)} 開始・還沒按完成`, by && `${by} 點`].filter(Boolean).join('・')
  const edited = p.editedAt ? `${fmtTime(p.editedAt)} ${p.editedBy ? `${p.editedBy} ` : ''}改過` : ''
  const tile = (n, label, cls) => `<div class="pk-stat ${cls}${n ? '' : ' zero'}"><b>${n}</b><span>${label}</span></div>`
  const probRow = ({ l, info, st }) => {
    const rt = readText(l)
    const meta =
      st === 'none' ? `單子上寫：${rt || info.name}・單子 ${fmtQty(l.qty) || '？'}` : st === 'guess' ? `AI 配的，沒確認・單子上寫：${rt || '—'}` : [info.code, `單子 ${fmtQty(l.qty) || '？'}`, `實拿 ${l.done ? fmtQty(gotOf(l)) || '—' : '—'}`].filter(Boolean).join('・')
    return `<div class="row pk-rrow st-${st}">${pickThumb(st === 'none' ? { it: null, name: rt || info.name } : info, 40)}<span class="grow"><span class="title">${esc(st === 'none' ? rt || info.name : info.name)}</span><br><span class="meta">${esc(meta)}</span></span>${stateChip(l, info.unit)}</div>`
  }
  const okRow = ({ l, info }) => `<div class="row pk-rrow st-ok">${pickThumb(info, 36)}<span class="grow"><span class="title">${esc(info.name)}</span>${info.code ? `<br><span class="meta">${esc(info.code)}</span>` : ''}</span><span class="qty"><b>${esc(fmtQty(gotOf(l)))}</b><small>${esc(info.unit || '個')}</small></span></div>`
  const verdict = !s.total ? '沒有品項' : allOk ? `全部對了・${s.total} 項` : `${probs.length} 項要處理・共 ${s.total} 項`
  return `
  <main class="app">
    <div class="nav">${backBtn('pick', '點貨')}<span class="nav-right">${canEdit() && p.doneAt ? '<button class="btn small secondary" data-action="pk-reopen">改一下</button>' : ''}${canManage() && p.doneAt ? `<button class="icon-btn" data-action="pk-delete" aria-label="刪除這筆點貨紀錄" title="刪除這筆點貨紀錄">${icon('trash', 22)}</button>` : ''}</span></div>
    <div class="pk-result">
      <div class="pk-res-side">
        <section class="pk-res-card" aria-label="點貨結果">
          <div class="pk-res-kind"><span class="pk-kind-ico ${kind}" aria-hidden="true">${icon(kind, 22)}</span><span class="grow"><b>${PICK_TITLE[kind]}結果</b> ${betaBadge}<br><span class="meta">${esc(head)}</span>${edited ? `<br><span class="meta pk-edited">${icon('edit', 14)}${esc(edited)}</span>` : ''}</span></div>
          <h1 class="pk-verdict ${allOk ? 'ok' : 'bad'}">${icon(allOk ? 'check' : 'warning', 28)}<span>${esc(verdict)}</span></h1>
          <div class="pk-stats">${tile(s.ok, '對了', 'ok')}${tile(s.bad, '數量不對', 'bad')}${tile(s.todo, '還沒點', 'todo')}${tile(s.none, '對不到', 'none')}</div>
        </section>
        <div class="pk-res-actions">
          <button class="btn secondary block" data-action="pk-copy">${icon('copy', 20)}複製結果（貼到 LINE）</button>
          <p class="footnote">點貨不會改帳面（帳面以正航為準），只留這筆紀錄${syncReady() ? '，大家的手機都看得到' : ''}。要傳給老闆：直接截圖，或按上面複製文字。</p>
        </div>
      </div>
      <div class="pk-res-lists">
        ${probs.length ? `<p class="section-title">要處理的（${probs.length}）</p><div class="group">${probs.map(probRow).join('')}</div>` : ''}
        ${oks.length ? `<p class="section-title">對了（${oks.length}）</p><div class="group${oks.length >= 6 ? ' cols-2' : ''}">${oks.map(okRow).join('')}</div>` : ''}
      </div>
    </div>
  </main>`
}

// —— 數量面板（可以有小數；打 0 或看不懂不會關掉，直接在面板裡說） ——
function qtySheet({ title, sub = '', value = '', action = '儲存', quick = [], skip = '' }, onSave, onSkip) {
  sheet(
    `<h2 class="sheet-title">${esc(title)}</h2>${sub ? `<p class="sheet-sub">${sub}</p>` : ''}
     <input class="field big-num" id="q-val" inputmode="decimal" autocomplete="off" value="${esc(value)}" aria-label="${esc(title)}" aria-describedby="q-err">
     <p class="q-err" id="q-err" role="alert" hidden></p>
     ${quick.length ? `<div class="q-quick">${quick.map((n) => `<button class="chip" data-n="${n}">${n}</button>`).join('')}</div>` : ''}
     <div class="row-actions" style="margin-top:14px"><button class="btn secondary" id="q-cancel" style="flex:1">${esc(skip || '取消')}</button><button class="btn" id="q-save" style="flex:1">${esc(action)}</button></div>`,
    (el, close) => {
      const input = el.querySelector('#q-val')
      const err = el.querySelector('#q-err')
      input.focus()
      input.select()
      const done = async (v) => {
        const n = parseQty(v)
        if (n == null || n <= 0) {
          err.hidden = false
          err.textContent = n === 0 ? '數量要大於 0' : '請打數字（可以有小數，最多兩位）'
          input.focus()
          return
        }
        close()
        await onSave(n)
      }
      el.querySelector('#q-cancel').onclick = () => {
        close()
        onSkip?.()
      }
      el.querySelector('#q-save').onclick = () => done(input.value)
      input.addEventListener('keydown', (e) => e.key === 'Enter' && done(input.value))
      input.addEventListener('input', () => (err.hidden = true))
      el.querySelectorAll('[data-n]').forEach((b) => (b.onclick = () => done(b.dataset.n)))
    },
    () => onSkip?.(),
  )
}

// —— 動作 ——
const sameProduct = (a, b) => !!((a.no && a.no === b.no) || (a.itemId && a.itemId === b.itemId))
/**
 * 加一批品項（AI 讀到的、搜尋加的）：確定是同一個產品、已經在清單裡（或這一批裡重複）的，一次問整批：略過，還是把數量加上去。
 * AI 猜的（guess）不算重複：產品還不確定，按「是這個」時才問。
 */
function addLines(p, lines, { quiet = false } = {}) {
  const fresh = []
  const dups = []
  for (const l of lines) {
    const target = !l.guess && (l.no || l.itemId) ? [...p.lines, ...fresh].find((x) => !x.guess && sameProduct(x, l)) : null
    if (target) dups.push({ l, target })
    else fresh.push(l)
  }
  const commit = async (mode) => {
    if (p._gone) return
    p.lines.push(...fresh)
    // 數量加上去：單子變多了，那一項要重新點（不然會直接變紅「少 N」）
    if (mode === 'add') for (const { l, target } of dups) if (l.qty != null) Object.assign(target, { qty: round2((target.qty || 0) + l.qty), done: false })
    await pickSave(p)
    if (!quiet) {
      const miss = fresh.filter((l) => !l.no && !l.itemId).length
      const guess = fresh.filter((l) => l.guess).length
      const merged = mode === 'add' ? dups.length : 0
      const skipped = mode === 'skip' ? dups.length : 0
      const parts = [`加了 ${fresh.length} 項`, merged && `${merged} 項數量加到原本那一項`, skipped && `略過 ${skipped} 項重複的`, guess && `${guess} 項是 AI 配的、請確認`, miss && `${miss} 項對不到產品`, fresh.some((l) => l.qty == null) && '有數量沒讀到的，點「單子」那格補上'].filter(Boolean)
      toast(parts.join('；'))
    }
    if (state.pick === p && state.view === 'pick-edit') render()
  }
  if (!dups.length) return commit()
  const where = (t) => {
    const i = p.lines.indexOf(t)
    return i >= 0 ? `第 ${i + 1} 項` : '這次讀到的另一列'
  }
  const ix = pickIxNow()
  sheet(
    `<h2 class="sheet-title">${dups.length === 1 ? `已經在${where(dups[0].target)}` : `有 ${dups.length} 項已經在清單裡`}</h2>
     <p class="sheet-sub">同一個產品不要重複點：要略過，還是把數量加上去？</p>
     <div class="group">${dups
       .slice(0, 8)
       .map(({ l, target }) => `<div class="row"><span class="grow"><span class="title pk-tname">${esc(lineInfo(ix, target).name)}</span><br><span class="meta">${esc(where(target))}：單子 ${esc(fmtQty(target.qty) || '？')}・這次 ${esc(fmtQty(l.qty) || '？')}</span></span></div>`)
       .join('')}${dups.length > 8 ? `<div class="row muted">還有 ${dups.length - 8} 項</div>` : ''}</div>
     <div class="row-actions" style="margin-top:14px"><button class="btn secondary" id="dp-skip" style="flex:1">略過</button><button class="btn" id="dp-add" style="flex:1">數量加上去</button></div>`,
    (el, close) => {
      el.querySelector('#dp-skip').onclick = () => {
        close()
        commit('skip')
      }
      el.querySelector('#dp-add').onclick = () => {
        close()
        commit('add')
      }
    },
    () => commit('skip'),
  )
}
/** 從搜尋結果加一項：已經在清單裡 → 問要不要把數量加上去；沒有 → 問單子上寫幾個。加完留在搜尋框，新的那張卡片閃一下 */
function addPickLine(r) {
  const p = state.pick
  if (!p || !r) return
  const t = pickTarget(r)
  const target = p.lines.find((l) => !l.guess && sameProduct(l, t))
  const after = (id) => {
    state.pickQ = ''
    state.pickFlash = { id, scroll: false }
    state.pickFocusQ = true
    render()
  }
  if (target)
    return qtySheet({ title: `已經在第 ${p.lines.indexOf(target) + 1} 項`, sub: `${esc(t.name)}：單子現在 ${esc(fmtQty(target.qty) || '？')}。要把數量加上去嗎？加幾個？`, action: '數量加上去', skip: '略過', quick: [1, 2, 3, 5, 10] }, async (n) => {
      target.qty = round2((target.qty || 0) + n)
      target.done = false // 單子變多了，要重新點
      await pickSave(p)
      toast(`第 ${p.lines.indexOf(target) + 1} 項的單子數量改成 ${fmtQty(target.qty)}`)
      after(target.id)
    })
  if (p.lines.length >= PICK_MAX_LINES) return toast(`一張最多 ${PICK_MAX_LINES} 項`)
  qtySheet({ title: '單子上寫幾個？', sub: esc(t.name), action: '加進清單', quick: [1, 2, 3, 5, 10] }, async (n) => {
    const l = { id: uid(), ...t, qty: n, got: null, done: false }
    p.lines.push(l)
    await pickSave(p)
    toast(`已加入：${t.name}`)
    after(l.id)
  })
}
function editLineQty(l, then) {
  qtySheet({ title: '單子上寫幾個？', sub: esc(lineInfo(pickIxNow(), l).name), value: fmtQty(l.qty), action: '儲存' }, async (n) => {
    l.qty = n
    then?.()
    await pickSave()
    render()
  })
}
/** 這一項跟清單裡另一項是同一個產品：把數量加過去（合併成一項），或刪掉這一項 */
function mergeIntoSheet(p, l, target) {
  const n = p.lines.indexOf(target) + 1
  sheet(
    `<h2 class="sheet-title">已經在第 ${n} 項</h2>
     <p class="sheet-sub">${esc(lineInfo(pickIxNow(), target).name)}：第 ${n} 項單子 ${esc(fmtQty(target.qty) || '？')}、這一項單子 ${esc(fmtQty(l.qty) || '？')}。要略過（刪掉這一項），還是把數量加到第 ${n} 項？</p>
     <div class="row-actions" style="margin-top:14px"><button class="btn secondary" id="mg-skip" style="flex:1">略過</button><button class="btn" id="mg-add" style="flex:1">數量加上去</button></div>`,
    (el, close) => {
      const finish = (add) => {
        close()
        if (add && l.qty != null) Object.assign(target, { qty: round2((target.qty || 0) + l.qty), done: false })
        const i = p.lines.indexOf(l)
        if (i >= 0) p.lines.splice(i, 1)
        pickSave(p)
        state.pickFlash = { id: target.id, scroll: true }
        render()
        toast(add ? `已合併到第 ${p.lines.indexOf(target) + 1} 項：單子 ${fmtQty(target.qty)}` : '已略過（刪掉這一項）')
      }
      el.querySelector('#mg-skip').onclick = () => finish(false)
      el.querySelector('#mg-add').onclick = () => finish(true)
    },
  )
}
/** 選產品（對不到的、AI 配的換一個、對錯的）：搜尋框先帶入 AI 讀到的編號或品名 */
async function assignLine(l) {
  const ix = await pickCtx()
  const rt = readText(l)
  const q0 = (l.no || l.itemId ? '' : l.read?.code || l.read?.name) || ''
  sheet(
    `<h2 class="sheet-title">這是哪一個產品？</h2>
     <p class="sheet-sub">${rt ? `單子上寫：「${esc(rt)}」。` : `現在是：${esc(lineInfo(ix, l).name)}。`}打產品編號，或品名、型號找。</p>
     <input class="field search" id="pp-q" type="search" placeholder="打產品編號，或品名、型號" autocomplete="off" spellcheck="false" enterkeyhint="search" value="${esc(q0)}">
     <div id="pp-results" style="margin-top:10px">${pickResultsHtml(ix, q0, 'assign')}</div>`,
    (el, close) => {
      const input = el.querySelector('#pp-q')
      const box = el.querySelector('#pp-results')
      let timer = 0
      input.addEventListener('input', () => {
        clearTimeout(timer)
        timer = setTimeout(() => {
          box.innerHTML = pickResultsHtml(ix, input.value, 'assign')
          glueTails(box)
        }, 120)
      })
      box.addEventListener('click', (e) => {
        const b = e.target.closest('[data-pp]')
        if (!b) return
        const r = b.dataset.no ? rowOfNo(ix, b.dataset.no) : ix.itemById.get(b.dataset.iid) ? { it: ix.itemById.get(b.dataset.iid) } : null
        if (!r) return
        close()
        const p = state.pick
        const t = pickTarget(r)
        delete l.no
        delete l.itemId
        delete l.guess
        delete l.cands
        Object.assign(l, t)
        // 選的產品清單裡已經有了：問要合併還是略過
        const target = p?.lines.find((x) => x !== l && !x.guess && sameProduct(x, l))
        if (target) return mergeIntoSheet(p, l, target)
        pickSave()
        render()
        toast(`已對到：${t.name}`)
      })
      if (!q0) input.focus()
    },
  )
}
/** AI 配的 → 按「是這個」才算確認（清單裡已經有同一個產品就問要不要合併） */
function confirmLine(l) {
  const p = state.pick
  delete l.guess
  const target = p?.lines.find((x) => x !== l && !x.guess && sameProduct(x, l))
  if (target) return mergeIntoSheet(p, l, target)
  pickSave()
  render()
}
function removeLine(l) {
  const p = state.pick
  const i = p?.lines.indexOf(l) ?? -1
  if (i < 0) return
  p.lines.splice(i, 1)
  pickSave(p)
  render()
  toast('已刪掉這一項', {
    label: '復原',
    run: () => {
      if (state.pick !== p) return toast('已經換到別張了，沒辦法復原')
      p.lines.splice(Math.min(i, p.lines.length), 0, l)
      pickSave(p)
      render()
    },
  })
}
function lineMoreSheet(l) {
  const info = lineInfo(pickIxNow(), l)
  const amber = ['none', 'guess'].includes(lineState(l))
  sheet(
    `<h2 class="sheet-title">${esc(amber ? readText(l) || info.name : info.name)}</h2>${info.code && !amber ? `<p class="sheet-sub">${esc(info.code)}</p>` : ''}
     <div class="group">
       <button class="row" id="lm-qty"><span class="grow"><span class="title">改單子數量</span><br><span class="meta">現在：${esc(fmtQty(l.qty) || '沒填')}</span></span>${chev}</button>
       <button class="row" id="lm-swap"><span class="grow"><span class="title">${amber ? '選產品' : '換產品'}</span><br><span class="meta">${amber ? '選是哪一個產品' : '對錯產品時用'}</span></span>${chev}</button>
       ${l.done ? `<button class="row" id="lm-undo"><span class="grow"><span class="title">改回還沒點</span><br><span class="meta">拿錯了、要重點的時候用</span></span>${chev}</button>` : ''}
       <button class="row" id="lm-del"><span class="grow"><span class="title" style="color:var(--red-ink)">刪掉這項</span><br><span class="meta">刪掉後幾秒內可以復原</span></span>${chev}</button>
     </div>`,
    (el, close) => {
      el.querySelector('#lm-qty').onclick = () => {
        close()
        editLineQty(l)
      }
      el.querySelector('#lm-swap').onclick = () => {
        close()
        assignLine(l)
      }
      el.querySelector('#lm-undo')?.addEventListener('click', () => {
        close()
        l.done = false
        paintLine(l)
        pickSave()
        toast('已改回還沒點', {
          label: '復原',
          run: () => {
            l.done = true
            paintLine(l)
            pickSave()
          },
        })
      })
      el.querySelector('#lm-del').onclick = () => {
        close()
        removeLine(l)
      }
    },
  )
}
/** 按「完成」：還有還沒點的先提醒（回去點／還是完成） */
function finishPickSheet(p) {
  const ix = pickIxNow()
  const undone = p.lines.filter((l) => !l.done || ['none', 'guess'].includes(lineState(l)))
  const amber = undone.filter((l) => ['none', 'guess'].includes(lineState(l))).length
  sheet(
    `<h2 class="sheet-title">還有 ${undone.length} 項還沒點</h2>
     <p class="sheet-sub">還是完成的話，這些會記成「還沒點」${amber ? `（其中 ${amber} 項還沒確認是哪一個產品，記成「對不到」）` : ''}。</p>
     <div class="group">${undone
       .slice(0, 6)
       .map((l) => `<div class="row"><span class="grow"><span class="title pk-tname">${esc(lineInfo(ix, l).name)}</span><br><span class="meta">單子 ${esc(fmtQty(l.qty) || '？')}</span></span>${stateChip(l, lineInfo(ix, l).unit)}</div>`)
       .join('')}${undone.length > 6 ? `<div class="row muted">還有 ${undone.length - 6} 項</div>` : ''}</div>
     <div class="row-actions" style="margin-top:14px"><button class="btn secondary" id="pf-back" style="flex:1">回去點</button><button class="btn" id="pf-go" style="flex:1">還是完成</button></div>`,
    (el, close) => {
      el.querySelector('#pf-back').onclick = () => {
        close()
        flashLine(undone[0].id)
      }
      el.querySelector('#pf-go').onclick = () => {
        close()
        finishPick(p)
      }
    },
  )
}
/**
 * 完成：
 * - 第一次完成：記完成時間、存進紀錄、同步
 * - 「改一下」完成修改：至少要留 1 項；雲端在這段時間被別台改過 → 用每一列的 id 合併兩邊；記誰、幾點改的
 */
async function finishPick(p) {
  if (p._gone) return toast('這張在另一台被刪掉了')
  if (state.pickBusyId === p.id) return toast('AI 還在讀單子，讀完再按完成')
  const now = Date.now()
  const copy = editingCopy() && state.pick === p
  let out = p
  let merged = false
  if (copy) {
    if (!p.lines.length) return toast('完成的紀錄至少要留 1 項；整張都不要了，請擁有者或管理員在結果頁刪除')
    const base = state.pickEdit.base
    const latest = await idb.picks.get(p.id).catch(() => null)
    if (!latest) return toast('這張在另一台被刪掉了，沒辦法存')
    const { _copy, ...mine } = p
    merged = (latest.updatedAt || 0) !== (base.updatedAt || 0)
    out = merged ? mergePick(base, mine, latest) : mine
    out.doneAt = base.doneAt || now
    out.editedAt = now
    out.editedBy = whoAmI()
  } else out.doneAt = out.doneAt || now
  // 每一項的狀態也記一份（ok 對了／bad 數量不對／todo 還沒點／none 對不到／guess AI 配的沒確認）：紀錄自己就看得懂
  for (const l of out.lines) l.st = lineState(l)
  if (pickSaving) await pickSaving
  try {
    await savePick(out)
  } catch {
    if (!copy) delete out.doneAt
    return
  }
  state.pick = out
  state.pickEdit = null
  state.pickQ = ''
  go('pick-result')
  toast(copy ? (merged ? '改好了：另一台也改了這張，已經把兩邊的修改合在一起' : '改好了：已更新這筆點貨紀錄') : '點完了：已存成點貨紀錄（不會改帳面）')
  if (syncReady()) syncNow().catch((err) => toast(`同步沒成功：${err.message}；有網路時會再自動試`))
}
/** 複製結果（貼到 LINE 給老闆）：不寫客戶、價格（本來就沒存） */
function pickText(p, ix) {
  const s = pickSummary(p)
  const rows = p.lines.map((l) => ({ l, info: lineInfo(ix, l), st: lineState(l) }))
  const one = ({ l, info, st }) => `・${st === 'none' ? readText(l) || info.name : info.name}${info.code && st !== 'none' ? `（${info.code}）` : ''}：單子 ${fmtQty(l.qty) || '？'}，實拿 ${l.done ? fmtQty(gotOf(l)) || '—' : '—'}`
  const sec = (title, sts, extra = () => '') => {
    const xs = rows.filter((x) => sts.includes(x.st))
    return xs.length ? [`【${title}】`, ...xs.map((x) => one(x) + extra(x))] : []
  }
  return [
    `${PICK_TITLE[p.kind] || '點貨'}結果（測試版）`,
    `單號：${p.ref || '沒填'}`,
    `時間：${fmtTime(p.doneAt || p.createdAt)}${byName(p) ? `　點貨人：${byName(p)}` : ''}`,
    ...(p.editedAt ? [`改過：${fmtTime(p.editedAt)}${p.editedBy ? ` ${p.editedBy}` : ''}`] : []),
    `對了 ${s.ok} 項・數量不對 ${s.bad} 項・還沒點 ${s.todo} 項・對不到 ${s.none} 項`,
    ...sec('數量不對', ['bad'], ({ l, info }) => `（${l.qty == null ? '單子數量沒填' : diffText(l, info.unit)}）`),
    ...sec('還沒點', ['todo']),
    ...sec('對不到產品', ['none', 'guess'], ({ st }) => (st === 'guess' ? '（AI 配的，沒確認）' : '')),
    ...sec('對了', ['ok']),
  ].join('\n')
}

// —— 拍單子讓 AI 讀：先框出品項那一段、遮住單價和金額，只送框裡、沒遮住的部分 ——
const NOTE_SCHEMA = {
  type: 'OBJECT',
  properties: {
    rows: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          code: { type: 'STRING', description: '產品編號（料號），照單子原樣；沒有就空字串' },
          name: { type: 'STRING', description: '品名，照單子原樣' },
          spec: { type: 'STRING', description: '規格：型號、尺寸；沒有就空字串' },
          qty: { type: 'NUMBER', description: '數量；可以有小數（例如 2.5）；看不清楚填 0' },
        },
        required: ['code', 'name', 'spec', 'qty'],
      },
    },
  },
  required: ['rows'],
}
const NOTE_PROMPT = `照片是一張出貨單或進貨單裡「品項」那一段（上面的客戶資料、下面的金額合計已經裁掉；黑色長條是遮住的欄位，不用讀）。
請一列一列讀出每一個品項：code＝產品編號、name＝品名、spec＝規格（型號、尺寸）、qty＝數量。
規則：
1. 只讀品項的產品編號、品名、規格、數量。客戶名稱、地址、電話、統一編號、單價、金額、小計、合計、稅額、備註一律忽略，不要寫進任何欄位。
2. 照單子上的字原樣抄；看不清楚的欄位留空字串，不要猜。
3. qty 只填數量那一欄的數字（可以有小數，例如冷媒 2.5）；看不清楚填 0。
4. 品名和規格寫在同一格的：品名放 name，型號、尺寸放 spec。
5. 標題列、空白列、合計列不是品項，不要列出來。
只回 JSON。`
/** 預設的框：單子中間的品項區（上面約 22% 是公司、客戶資料，下面約 18% 是金額合計、簽名） */
const CROP_DEFAULT = { x1: 0.03, y1: 0.22, x2: 0.97, y2: 0.82 }
const CROP_MIN = 0.08
const MASK_MIN = 0.02
function cropRect() {
  const r = readJson(LS.pickCrop, null)
  const ok = r && ['x1', 'y1', 'x2', 'y2'].every((k) => Number.isFinite(r[k])) && r.x1 >= 0 && r.y1 >= 0 && r.x2 <= 1 && r.y2 <= 1 && r.x2 - r.x1 >= CROP_MIN && r.y2 - r.y1 >= CROP_MIN
  return ok ? { x1: r.x1, y1: r.y1, x2: r.x2, y2: r.y2 } : { ...CROP_DEFAULT }
}
/** 遮住的欄位（直的長條，左右位置用照片寬度的比例）：跟框一樣記在這台 */
function cropMasks() {
  const list = readJson(LS.pickMask, [])
  return (Array.isArray(list) ? list : [])
    .filter((m) => m && Number.isFinite(m.x1) && Number.isFinite(m.x2) && m.x1 >= 0 && m.x2 <= 1 && m.x2 - m.x1 >= MASK_MIN)
    .slice(0, PICK_MASK_MAX)
    .map((m) => ({ x1: m.x1, x2: m.x2 }))
}
const CROP_HANDLES = { nw: '拖曳左上角', ne: '拖曳右上角', sw: '拖曳左下角', se: '拖曳右下角', n: '拖曳上邊', s: '拖曳下邊', w: '拖曳左邊', e: '拖曳右邊' }
/** 第一次用的示意圖（自己畫的，不是真的單子）：黃框框住品項那幾列、黑色長條遮住單價和金額 */
const CROP_HELP_SVG = `<svg class="crop-help-svg" viewBox="0 0 300 196" role="img" aria-label="示意圖：黃框框住品項那幾列，黑色長條遮住單價、金額兩欄">
  <rect x="58" y="4" width="184" height="188" rx="6" fill="#ffffff" stroke="#c7c7cc"/>
  <rect x="70" y="14" width="160" height="28" rx="3" fill="#f6dada"/><text x="150" y="33" text-anchor="middle" font-size="11" fill="#a3191f">客戶名稱、地址、電話</text>
  <g font-size="10" fill="#3a3a3c"><text x="78" y="60">品名規格</text><text x="148" y="60">數量</text><text x="178" y="60">單價</text><text x="206" y="60">金額</text></g>
  <g fill="#8e8e93">${[72, 88, 104, 120, 136].map((y) => `<rect x="78" y="${y}" width="52" height="5" rx="2"/><rect x="150" y="${y}" width="12" height="5" rx="2"/><rect x="178" y="${y}" width="18" height="5" rx="2"/><rect x="206" y="${y}" width="22" height="5" rx="2"/>`).join('')}</g>
  <rect x="70" y="154" width="160" height="28" rx="3" fill="#dde3f6"/><text x="150" y="172" text-anchor="middle" font-size="11" fill="#1f44a8">合計、稅額、簽名</text>
  <rect x="172" y="48" width="58" height="98" fill="#000000" fill-opacity="0.85"/><text x="201" y="101" text-anchor="middle" font-size="11" fill="#ffffff">遮住</text>
  <rect x="72" y="48" width="158" height="98" fill="none" stroke="#ffd60a" stroke-width="3"/>
  <g font-size="10" fill="#ffffff"><text x="52" y="100" text-anchor="end">黃框</text><text x="52" y="113" text-anchor="end">只框品項</text></g>
  <path d="M54 104h14" stroke="#ffd60a" stroke-width="2"/>
</svg>`
/**
 * 框出品項、遮住單價和金額：照片上自動框好一個長方形（上次框的位置，記在這台），
 * 四個角、四個邊可以拖（手指範圍 44px），框裡面拖＝整個移動；按「遮住欄位」後，在框裡左右拖＝畫一條要塗黑的直條。
 * 回傳 { rect 框（0～1 的比例）, masks 遮住的直條 }，取消回傳 null。
 */
function cropSheet(blob) {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(blob)
    const r = cropRect()
    let masks = cropMasks()
    let maskMode = false
    const el = document.createElement('div')
    el.className = 'crop-wrap'
    el.setAttribute('role', 'dialog')
    el.setAttribute('aria-modal', 'true')
    el.setAttribute('aria-label', '框出單子上的品項')
    el.innerHTML = `<div class="crop-bar"><button type="button" class="btn small crop-plain" data-c="cancel">取消</button><b class="crop-title">框出品項、遮住金額</b><button type="button" class="btn small crop-plain" data-c="help">怎麼框？</button></div>
      <p class="crop-tip" id="crop-tip"></p>
      <div class="crop-stage"><div class="crop-img"><img src="${url}" alt="拍到的單子" draggable="false"><div class="crop-dim" aria-hidden="true"><div class="crop-hole"></div></div><div class="crop-area" data-h="move"></div><div class="crop-masks"></div><div class="crop-rect">${Object.entries(CROP_HANDLES)
        .map(([h, label]) => `<button type="button" class="crop-h ${h}" data-h="${h}" aria-label="${label}（可以用方向鍵）"><span></span></button>`)
        .join('')}</div></div></div>
      <div class="crop-foot"><div class="crop-tools"><button type="button" class="btn small crop-plain" data-c="mask" aria-pressed="false">遮住欄位</button><button type="button" class="btn small crop-plain" data-c="clear">清掉遮罩</button><button type="button" class="btn small crop-plain" data-c="reset">重設框</button></div>
      <p class="crop-note" id="crop-note"></p><button type="button" class="btn block" data-c="send">送出，請 AI 讀</button></div>
      <div class="crop-help" hidden><div class="crop-help-card" role="document"><b class="crop-help-title">先框、再遮</b>${CROP_HELP_SVG}<ol><li><b>黃框</b>只框住品項那幾列：上面的客戶、下面的合計框在外面。</li><li>單價、金額跟品項在同一列：按「<b>遮住欄位</b>」，在框裡左右拖，把那幾欄塗黑。</li><li>框和遮的位置記在這台，下次自動套用；看一眼按「送出」就好。</li></ol><button type="button" class="btn block" data-c="help-ok">知道了</button></div></div>`
    document.body.append(el)
    document.body.classList.add('no-scroll')
    const box = el.querySelector('.crop-rect')
    const area = el.querySelector('.crop-area')
    const hole = el.querySelector('.crop-hole')
    const maskBox = el.querySelector('.crop-masks')
    const imgBox = el.querySelector('.crop-img')
    const help = el.querySelector('.crop-help')
    const tip = el.querySelector('#crop-tip')
    const note = el.querySelector('#crop-note')
    const pos = (n, a) => Object.assign(n.style, { left: `${a.x1 * 100}%`, top: `${a.y1 * 100}%`, width: `${(a.x2 - a.x1) * 100}%`, height: `${(a.y2 - a.y1) * 100}%` })
    const paintMasks = () => [...maskBox.children].forEach((n, i) => masks[i] && pos(n, { x1: masks[i].x1, x2: masks[i].x2, y1: r.y1, y2: r.y2 }))
    const renderMasks = () => {
      maskBox.innerHTML = masks.map((_, i) => `<div class="crop-mask" data-m="${i}"><button type="button" class="crop-mh l" data-mh="l" data-m="${i}" aria-label="拖曳遮罩的左邊（可以用方向鍵）"><span></span></button><span class="crop-mlabel">遮住</span><button type="button" class="crop-mx" data-mx="${i}" aria-label="拿掉這個遮罩">${icon('x', 16)}</button><button type="button" class="crop-mh r" data-mh="r" data-m="${i}" aria-label="拖曳遮罩的右邊（可以用方向鍵）"><span></span></button></div>`).join('')
      paintMasks()
      words()
    }
    const words = () => {
      tip.textContent = maskMode ? '在黃框裡左右拖，畫出要塗黑的欄位（單價、金額）；黑條的兩邊可以拉、× 可以拿掉。' : '拖曳黃框的四個角或四個邊，只框住品項那幾列。'
      note.innerHTML = `${icon('shield', 18)}<span>${masks.length ? '只會送框裡、沒遮住的部分給 AI。' : '只會送框裡的部分給 AI。單價、金額在框裡的話，請按「遮住欄位」塗掉。'}</span>`
      const mb = el.querySelector('[data-c="mask"]')
      mb.setAttribute('aria-pressed', String(maskMode))
      mb.classList.toggle('on', maskMode)
      area.classList.toggle('drawing', maskMode)
      el.querySelector('[data-c="clear"]').disabled = !masks.length
    }
    const paint = () => {
      for (const n of [box, hole, area]) pos(n, r)
      paintMasks()
    }
    paint()
    renderMasks()
    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v))
    const apply = (h, s, dx, dy) => {
      let { x1, y1, x2, y2 } = s
      if (h === 'move') {
        const w = x2 - x1
        const hh = y2 - y1
        x1 = clamp(x1 + dx, 0, 1 - w)
        y1 = clamp(y1 + dy, 0, 1 - hh)
        x2 = x1 + w
        y2 = y1 + hh
      } else {
        if (h.includes('n')) y1 = clamp(y1 + dy, 0, y2 - CROP_MIN)
        if (h.includes('s')) y2 = clamp(y2 + dy, y1 + CROP_MIN, 1)
        if (h.includes('w')) x1 = clamp(x1 + dx, 0, x2 - CROP_MIN)
        if (h.includes('e')) x2 = clamp(x2 + dx, x1 + CROP_MIN, 1)
      }
      Object.assign(r, { x1, y1, x2, y2 })
      paint()
    }
    const moveMask = (i, kind, s, dx) => {
      const m = masks[i]
      if (!m) return
      if (kind === 'move') {
        const w = s.x2 - s.x1
        m.x1 = clamp(s.x1 + dx, 0, 1 - w)
        m.x2 = m.x1 + w
      } else if (kind === 'l') m.x1 = clamp(s.x1 + dx, 0, s.x2 - MASK_MIN)
      else if (kind === 'r') m.x2 = clamp(s.x2 + dx, s.x1 + MASK_MIN, 1)
      paintMasks()
    }
    let drag = null
    el.addEventListener('pointerdown', (e) => {
      if (e.target.closest('[data-mx], .crop-help')) return
      const mh = e.target.closest('[data-mh]')
      const mk = e.target.closest('.crop-mask')
      const h = e.target.closest('[data-h]')
      if (!mh && !mk && !h) return
      e.preventDefault()
      const rect = imgBox.getBoundingClientRect()
      const base = { x: e.clientX, y: e.clientY, w: rect.width || 1, ht: rect.height || 1, id: e.pointerId }
      if (mh || mk) {
        const i = Number((mh || mk).dataset.m)
        drag = { ...base, type: 'mask', i, kind: mh ? mh.dataset.mh : 'move', s: { ...masks[i] } }
      } else if (h.dataset.h === 'move' && maskMode) {
        if (masks.length >= PICK_MASK_MAX) return toast(`最多遮 ${PICK_MASK_MAX} 欄`)
        const fx = clamp((e.clientX - rect.left) / base.w, 0, 1)
        masks.push({ x1: fx, x2: fx })
        renderMasks()
        drag = { ...base, type: 'new', i: masks.length - 1, fx }
      } else drag = { ...base, type: 'rect', h: h.dataset.h, s: { ...r } }
      try {
        e.target.setPointerCapture(e.pointerId)
      } catch {}
      box.classList.add('dragging')
    })
    el.addEventListener('pointermove', (e) => {
      if (!drag || e.pointerId !== drag.id) return
      const dx = (e.clientX - drag.x) / drag.w
      const dy = (e.clientY - drag.y) / drag.ht
      if (drag.type === 'rect') apply(drag.h, drag.s, dx, dy)
      else if (drag.type === 'mask') moveMask(drag.i, drag.kind, drag.s, dx)
      else if (drag.type === 'new') {
        const cur = clamp(drag.fx + dx, 0, 1)
        masks[drag.i] = { x1: Math.min(drag.fx, cur), x2: Math.max(drag.fx, cur) }
        paintMasks()
      }
    })
    const end = (e) => {
      if (!drag || e.pointerId !== drag.id) return
      // 只點一下、沒有拖：不要留下一條細細的遮罩
      if (drag.type === 'new' && masks[drag.i] && masks[drag.i].x2 - masks[drag.i].x1 < MASK_MIN) {
        masks.splice(drag.i, 1)
        renderMasks()
      } else words()
      drag = null
      box.classList.remove('dragging')
    }
    el.addEventListener('pointerup', end)
    el.addEventListener('pointercancel', end)
    const done = (v) => {
      el.remove()
      document.removeEventListener('keydown', onKey)
      document.body.classList.toggle('no-scroll', !!document.querySelector('.viewer'))
      URL.revokeObjectURL(url)
      resolve(v)
    }
    const showHelp = (on) => {
      help.hidden = !on
      if (on) help.querySelector('[data-c="help-ok"]').focus({ preventScroll: true })
      else el.querySelector('[data-c="send"]').focus({ preventScroll: true })
    }
    // 鍵盤也能調：選到某個角或遮罩的邊，按方向鍵移 1%（Shift 5%）；Esc 取消（說明打開時先關說明）
    const onKey = (e) => {
      if (e.key === 'Escape') return help.hidden ? done(null) : showHelp(false)
      const step = e.shiftKey ? 0.05 : 0.01
      const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key]
      if (!d) return
      const mh = e.target.closest?.('.crop-wrap [data-mh]')
      const t = e.target.closest?.('.crop-wrap [data-h]')
      if (mh) {
        e.preventDefault()
        const i = Number(mh.dataset.m)
        moveMask(i, mh.dataset.mh, { ...masks[i] }, d[0])
      } else if (t) {
        e.preventDefault()
        apply(t.dataset.h, { ...r }, d[0], d[1])
      }
    }
    document.addEventListener('keydown', onKey)
    el.addEventListener('click', (e) => {
      const mx = e.target.closest('[data-mx]')
      if (mx) {
        masks.splice(Number(mx.dataset.mx), 1)
        return renderMasks()
      }
      const c = e.target.closest('[data-c]')?.dataset.c
      if (c === 'cancel') done(null)
      else if (c === 'help') showHelp(true)
      else if (c === 'help-ok') {
        ls.set(LS.pickCropHelp, '1')
        showHelp(false)
      } else if (c === 'reset') {
        Object.assign(r, CROP_DEFAULT)
        paint()
      } else if (c === 'clear') {
        masks = []
        renderMasks()
      } else if (c === 'mask') {
        maskMode = !maskMode
        // 第一次按、還沒有遮罩：先在框的右邊放一條（單價、金額通常在右邊），拉一拉就好
        if (maskMode && !masks.length) {
          const w = r.x2 - r.x1
          masks.push({ x1: round3(r.x1 + w * 0.7), x2: round3(r.x2) })
          renderMasks()
        } else words()
      } else if (c === 'send') {
        const keep = Object.fromEntries(Object.entries(r).map(([k, v]) => [k, round3(v)]))
        const ms = masks.filter((m) => m.x2 - m.x1 >= MASK_MIN).map((m) => ({ x1: round3(m.x1), x2: round3(m.x2) }))
        ls.set(LS.pickCrop, JSON.stringify(keep))
        ls.set(LS.pickMask, JSON.stringify(ms))
        done({ rect: { ...r }, masks: ms })
      }
    })
    glueTails(el)
    // 第一次用：先看示意圖
    showHelp(ls.get(LS.pickCropHelp) !== '1')
  })
}
const round3 = (v) => Math.round(v * 1000) / 1000
/**
 * 在手機上把框裡那一塊切出來，遮住的欄位塗黑（canvas）：只有這一塊會送給 AI。
 * 從原始照片切（比先縮小再切清楚），切完再縮到長邊 1600。
 */
async function cropToBlob(src, r, masks = [], expectRatio = 0) {
  const bmp = await createImageBitmap(src, { imageOrientation: 'from-image' }).catch(() => bitmapOf(src))
  const W = bmp.width
  const H = bmp.height
  // 隱私保險（4.7.1 code review）：框和遮住的位置是畫在縮好的圖上；從原檔重新解碼時轉正方式如果不一樣，
  // 遮住的位置會跑掉、單價金額可能漏送 → 寬高比差超過 2% 就不用原檔（呼叫的地方會改從縮好的那張切）
  if (expectRatio && Math.abs(W / H / expectRatio - 1) > 0.02) {
    bmp.close?.()
    throw new Error('ratio mismatch')
  }
  const sx = Math.round(r.x1 * W)
  const sy = Math.round(r.y1 * H)
  const sw = Math.max(1, Math.min(W - sx, Math.round((r.x2 - r.x1) * W)))
  const sh = Math.max(1, Math.min(H - sy, Math.round((r.y2 - r.y1) * H)))
  const scale = Math.min(1, MAX_SIDE / Math.max(sw, sh))
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.round(sw * scale))
  canvas.height = Math.max(1, Math.round(sh * scale))
  const g = canvas.getContext('2d')
  g.drawImage(bmp, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height)
  bmp.close?.()
  g.fillStyle = '#000'
  for (const m of masks) {
    const a = Math.max(m.x1, r.x1)
    const b = Math.min(m.x2, r.x2)
    if (b <= a) continue
    // 往外多塗 2px：邊緣不會留下一條沒塗到的
    const x = Math.max(0, Math.floor((a - r.x1) * W * scale) - 2)
    g.fillRect(x, 0, Math.min(canvas.width, Math.ceil((b - r.x1) * W * scale) + 2) - x, canvas.height)
  }
  return new Promise((res) => canvas.toBlob(res, 'image/jpeg', 0.85))
}
/** AI 讀到的一列 → 點貨單的一項（確定的直接對上；AI 猜的標 guess；對不到的記有幾個像的） */
function lineFromRow(ix, row) {
  const m = matchRow(ix, row)
  const read = { code: row.code, name: row.name, spec: row.spec }
  if (m.r) return { id: uid(), ...pickTarget(m.r), qty: row.qty, got: null, done: false, ...(m.sure ? {} : { guess: true }), read }
  return { id: uid(), name: clip([row.name, row.spec].filter(Boolean).join(' ') || row.code, PICK_LEN.name), qty: row.qty, got: null, done: false, cands: m.cands, read }
}
/** 拍單子 → 框、遮 → 切 → AI 讀 → 每一列對到產品 */
async function readNoteFile(file) {
  const p = state.pick
  if (!p || !canEdit() || !hasAi()) return
  if (state.pickBusyId) return toast('AI 還在讀上一張，讀完再拍')
  if (p.lines.length >= PICK_MAX_LINES) return toast(`一張最多 ${PICK_MAX_LINES} 項`)
  let photo
  try {
    photo = await prepareImage(file)
  } catch {
    return toast('這張照片讀不進來，換一張試試')
  }
  const res = await cropSheet(photo.blob)
  if (!res) return
  const cropped = await cropToBlob(file, res.rect, res.masks, photo.w / photo.h).catch(() => cropToBlob(photo.blob, res.rect, res.masks).catch(() => null))
  if (!cropped) return toast('這張照片切不出來，換一張試試')
  sendNote(p, cropped)
}
/** 送給 AI 讀（失敗可以按「再試一次」：用剛剛切好的那張，不用重拍、重框） */
async function sendNote(p, cropped) {
  if (p._gone) return
  // 已經放棄的「改一下」副本：不要再叫 AI（白花額度，讀完也會丟掉）
  if (p._copy && state.pick !== p) return toast('這次修改已經放棄了')
  if (state.pickBusyId) return toast('AI 還在讀上一張，讀完再試')
  const doneAt0 = p.doneAt
  state.pickBusyId = p.id
  if (state.view === 'pick-edit') render()
  try {
    const res = await askJson([{ inline_data: { mime_type: 'image/jpeg', data: await blobToBase64(cropped) } }, { text: NOTE_PROMPT }], NOTE_SCHEMA, 40000)
    const rows = cleanNoteRows(res)
    // 讀的時候：這張被刪掉了、按完成了、或「改一下」已經放棄 → 讀到的不加進去
    if (p._gone) return
    if (p.doneAt !== doneAt0 || (p._copy && state.pick !== p)) return toast('這張已經不在編輯中，讀到的品項沒有加進去')
    if (!rows.length) return toast('AI 沒讀到品項：框要框住品項那幾列，再拍一次（拍正、不要反光）')
    const ix = await pickCtx()
    addLines(p, rows.slice(0, PICK_MAX_LINES - p.lines.length).map((row) => lineFromRow(ix, row)))
  } catch (e) {
    toast(e.message || 'AI 沒讀到', { label: '再試一次', run: () => sendNote(p, cropped) })
  } finally {
    if (state.pickBusyId === p.id) state.pickBusyId = null
    if (state.view === 'pick-edit') render()
  }
}
/** 匯入正航完成、交叉比對開著：跳出「這次有 N 項要查一下」，可以打開清單 */
function crossDoneSheet(msg, c) {
  const n = c.total ?? c.items?.length ?? 0
  const head = c.first ? '下次匯入才開始比對' : n ? `這次有 ${n} 項要查一下` : c.checked ? '這次沒有要查的' : '這次沒有可以比的'
  const sub = c.first
    ? '這次先記下每個產品的正航數量（基準）。之後用 App 點貨，下次匯入正航時就會比。'
    : n
      ? '上次匯入後有點過貨的產品裡，點貨記的進出跟正航的進出不一樣。'
      : c.checked
        ? `上次匯入後點過貨的 ${c.checked} 項，點貨跟正航的進出都一樣。`
        : '上次匯入後沒有點過貨的產品（或點貨紀錄還沒同步到這台）。上次的比對結果保留著。'
  sheet(
    `<h2 class="sheet-title">匯入完成</h2><p class="sheet-sub">${esc(msg)}</p>
     <div class="pk-cross-sum ${n ? 'bad' : 'ok'}"><span class="pk-cross-k">交叉比對（測試版） <span class="badge new">新</span></span><b>${esc(head)}</b><p>${esc(sub)}</p></div>
     <div class="row-actions" style="margin-top:14px">${n ? '<button class="btn secondary" id="cd-ok" style="flex:1">等一下再看</button><button class="btn" id="cd-open" style="flex:1">打開清單</button>' : '<button class="btn block" id="cd-ok">知道了</button>'}</div>`,
    (el, close) => {
      el.querySelector('#cd-ok').onclick = close
      el.querySelector('#cd-open')?.addEventListener('click', () => {
        close()
        crossListSheet(c)
      })
    },
  )
}
function crossListSheet(c) {
  const n = c.total ?? c.items?.length ?? 0
  const head = c.first ? '下次匯入才開始比對' : n ? `這次有 ${n} 項要查一下` : '這次沒有要查的'
  const rows = (c.items || [])
    .map(
      (x) => `<div class="row pk-cross-row"><span class="grow"><span class="title">${esc(x.name || x.no)}</span><br><span class="meta">${esc(x.no)}</span>
        <span class="pk-cross-nums"><span>點貨記的 <b>${esc(signed(x.pick))}</b></span><span>正航變了 <b>${esc(signed(x.erp))}</b></span></span>
        <span class="pk-cross-picks">${(x.picks || []).map((pk) => `<button class="chip" data-cross-pick="${esc(pk.id)}">${esc(PICK_KIND[pk.kind] || '點貨')} ${esc(fmtTime(pk.at))}${pk.ref ? `・${esc(pk.ref)}` : ''}・${esc(signed(pk.kind === 'in' ? pk.got : -pk.got))}</button>`).join('')}</span></span></div>`,
    )
    .join('')
  const synced = syncReady() ? `這台最後同步：${ls.get(LS.lastSync) ? fmtTime(Number(ls.get(LS.lastSync))) : '還沒同步過'}（別台還沒同步上來的點貨不會算進來）` : ''
  sheet(
    `<h2 class="sheet-title">交叉比對 ${betaBadge}</h2>
     <p class="sheet-sub">${esc(fmtTime(c.at))} 的正航報表・${esc(head)}${synced ? `<br>${esc(synced)}` : ''}</p>
     ${
       n
         ? `<div class="hint-card">不一定是誰錯：可能單子打錯、拿錯貨，或有進出貨沒用 App 點。</div><div class="group" style="margin-top:12px">${rows}</div>${n > (c.items || []).length ? `<p class="footnote">共 ${n} 項，這裡只列前 ${(c.items || []).length} 項。</p>` : ''}`
         : `<div class="group"><div class="row muted">${c.first ? '這次先記下基準（每個產品的正航數量）；下次匯入正航時，才會跟這段時間的點貨紀錄比。' : `上次匯入後點過貨的 ${c.checked} 項，點貨跟正航的進出都一樣。`}</div></div>`
     }
     <p class="footnote">怎麼算：上次匯入正航之後、這張報表之前按「完成」的點貨單，進貨實拿加起來、減掉出貨實拿＝點貨記的；這次正航的數量減掉上次的＝正航變了。兩個不一樣才列出來。${c.missing ? `另外有 ${c.missing} 項點過貨、但這次的正航表裡沒有，沒辦法比。` : ''}</p>
     <button class="btn block secondary" id="cx-ok" style="margin-top:12px">知道了</button>`,
    (el, close) => {
      el.querySelector('#cx-ok').onclick = close
      el.querySelectorAll('[data-cross-pick]').forEach(
        (b) =>
          (b.onclick = async () => {
            const p = await idb.picks.get(b.dataset.crossPick)
            if (!p) return toast('找不到這筆點貨紀錄（可能被刪掉了）')
            close()
            state.pick = p
            state.pickEdit = null
            go('pick-result')
          }),
      )
    },
  )
}
/** 別台傳來的點貨單：比這台新才用；正在「改一下」同一張 → 提示，按「完成修改」時合併 */
async function applyRemotePick(rec, cur, id) {
  const t = Math.min(Number(rec.t) || 0, Date.now() + 86400000)
  if (cur && tOf(cur) >= t) return false
  const d = normPick(rec.d, id)
  if (!d) return false
  // 這台有還沒上傳的修改（例如倉庫沒網路時改的），別台又比較晚改了同一張：不要整筆蓋掉，
  // 用每一列的 id 合在一起（這台改過的列用這台的、雲端才有的列加進來），再傳上去（4.7.1 code review）
  const localDirty = cur && cur.doneAt && tOf(cur) !== cur._syncT
  const v = localDirty ? { ...mergePick(null, cur, d), id, updatedAt: Math.max(Date.now(), t + 1), _syncT: undefined } : { ...d, updatedAt: t, _syncT: t }
  if (localDirty && !v.lines.length) return false
  await idb.picks.putRaw(v)
  if (localDirty) scheduleSync()
  picksCache = null
  if (state.pick?.id === id) {
    if (editingCopy()) {
      if (!state.pickEdit.warned) {
        state.pickEdit.warned = true
        toast('另一台也改了這張：按「完成修改」時，會把兩邊的修改合在一起')
      }
    } else if (state.view !== 'pick-edit') state.pick = v
  }
  return true
}
/** 點貨單那一頁的輸入框（render 之後綁） */
function bindPickInputs() {
  if (state.view !== 'pick-edit' || !state.pick) return
  const p = state.pick
  document.getElementById('pk-ref')?.addEventListener('input', (e) => {
    p.ref = clip(e.target.value, PICK_LEN.ref)
    pickSave(p)
  })
  const q = document.getElementById('pk-q')
  if (q) {
    let timer = 0
    q.addEventListener('input', () => {
      state.pickQ = q.value
      clearTimeout(timer)
      timer = setTimeout(async () => {
        const box = document.getElementById('pk-results')
        if (!box || state.pick !== p) return
        box.innerHTML = pickResultsHtml(await pickCtx(), q.value)
        glueTails(box)
      }, 120)
    })
    if (state.pickFocusQ) {
      state.pickFocusQ = false
      q.focus({ preventScroll: true })
    }
  }
  for (const id of ['pk-cam', 'pk-album'])
    document.getElementById(id)?.addEventListener('change', (e) => {
      const f = e.target.files?.[0]
      e.target.value = ''
      if (f) readNoteFile(f)
    })
  document.querySelectorAll('[data-pk-got]').forEach((input) => {
    const line = () => p.lines.find((x) => x.id === input.dataset.pkGot)
    input.addEventListener('input', () => {
      const l = line()
      if (!l) return
      const v = input.value.trim()
      // 打到一半（例如「2.」）還看不懂：先不改，等打完
      if (!v) l.got = null
      else {
        const n = parseQty(v)
        if (n == null) return
        l.got = n
      }
      paintLine(l)
      pickSave(p)
    })
    input.addEventListener('focus', () => input.select())
    // 離開輸入框：看不懂的字還原；已點好的空白＝跟單子一樣（顯示出來，結果才一致）
    input.addEventListener('change', () => {
      const l = line()
      if (!l) return
      if (input.value.trim() && parseQty(input.value) == null) input.value = fmtQty(l.got)
      if (l.done && l.got == null) l.got = l.qty
      input.value = fmtQty(l.got)
      paintLine(l)
      pickSave(p)
    })
  })
  if (state.pickFlash) {
    const { id, scroll } = state.pickFlash
    state.pickFlash = null
    requestAnimationFrame(() => flashLine(id, { scroll }))
  }
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
                      : state.view === 'catalog'
                        ? await viewCatalog()
                        : state.view === 'quality'
                          ? await viewQuality()
                          : state.view === 'pick'
                            ? await viewPick()
                            : state.view === 'pick-edit'
                              ? await viewPickEdit()
                              : state.view === 'pick-result'
                                ? await viewPickResult()
                                : await viewSettings()
  // 放大看照片時，重畫畫面不要讓位置跳回左上角
  const vs = $app.querySelector('.viewer-scroll')
  const keep = vs ? { x: vs.scrollLeft / Math.max(1, vs.scrollWidth), y: vs.scrollTop / Math.max(1, vs.scrollHeight) } : null
  // 平板、電腦：每一頁都有左邊側邊欄（子頁面在手機上不顯示）；body 記住現在哪一頁，給寬版排版用
  document.body.dataset.view = state.view
  $app.innerHTML = html.includes('class="tabbar') ? html : html + tabBar(TAB_OF[state.view] || 'home', true)
  const nv = $app.querySelector('.viewer-scroll')
  if (keep && nv) {
    nv.scrollLeft = keep.x * nv.scrollWidth
    nv.scrollTop = keep.y * nv.scrollHeight
  }
  if (nv) bindReviewPinch(nv)
  document.body.classList.toggle('no-scroll', !!document.querySelector('.viewer'))
  // 檢視者：所有修改用的按鈕藏起來（CSS：body.read-only .edit-only）
  document.body.classList.toggle('read-only', !canEdit())
  bindInputs()
  glueTails($app)
  // 液態玻璃鏡片：放在最後（量一次位置，不會讓瀏覽器多排一次版）
  mountLenses()
  // 新版等著裝：換到安全的頁面（例如回到首頁）就重新整理
  if (updateReady) setTimeout(tryReload, 300)
}

/**
 * 一個字單獨掉到下一行（「秒。」「選」「號）」）：把每一段說明最後 3 個字黏在一起，不在中間換行。
 * CSS 的 text-wrap: pretty 對中文常常沒效（iPhone 的 Safari 舊版也不支援），所以畫完再處理一次；不用量字的位置，很快。
 */
const GLUE_SEL = 'p, li, .meta, .footnote, .subtitle, .sheet-sub, .hint-card > div, .doubt-reason, .setup-card p, .pk-name, .pk-tname, .pk-rrow .title, .pk-res .title, .pk-cross-row .title'
function glueTails(root) {
  for (const el of root.querySelectorAll(GLUE_SEL)) {
    if ((el.textContent || '').length < 14) continue
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
    let last = null
    let n
    while ((n = walker.nextNode())) if (n.textContent.trim()) last = n
    if (!last || last.parentElement.closest('.nb, input, textarea, svg, script, style')) continue
    const t = last.textContent
    let k = t.length
    let count = 0
    while (k > 0 && count < 3) {
      k--
      if (!/\s/.test(t[k])) count++
    }
    if (k <= 0) continue // 這一小段本來就很短：不用黏
    const span = document.createElement('span')
    span.className = 'nb'
    span.textContent = t.slice(k)
    last.textContent = t.slice(0, k)
    last.after(span)
  }
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
  glueTails(back)
  return close
}

// ───────────────────────── 放大看照片（全站共用同一個） ─────────────────────────
// 10/10 使用者：「照片都應該要可以方便點開來看，讓使用者做二次確認。」
// 縮圖加上 zoomAttrs()：點一下（或 Tab 到它按 Enter）就用 photoViewer() 全螢幕看；不會觸發那一列原本的點擊。
// 結果頁有框的「放大看照片」（.viewer）用同一套黑底、右上 ×、Esc、兩指縮放（pinchZoom）。
const ZOOM_MAX = 4
const clampZoom = (z) => Math.max(1, Math.min(ZOOM_MAX, z))
/**
 * 縮圖的屬性字串：src 可以用換行分開放好幾張（例如一次盤點的所有照片）；
 * group：同一組的縮圖放大後可以左右換張（例如樣品照）；places：儲位代號（畫成黃標籤）；box：打開時直接放大到這一框。
 * inRow：縮圖包在 <button class="row"> 裡面 → 報讀器、Tab 只認那一列（不念兩次）；滑鼠、手指照樣點縮圖放大
 */
const zoomAttrs = (src, cap = '', { group = '', places = [], box = null, color = '', inRow = false } = {}) =>
  src
    ? ` data-zoom-src="${esc(src)}"${cap ? ` data-zoom-cap="${esc(cap)}"` : ''}${group ? ` data-zoom-group="${esc(group)}"` : ''}${places.length ? ` data-zoom-places="${esc(places.join('\n'))}"` : ''}${box ? ` data-zoom-box="${esc(box.join(','))}"` : ''}${color ? ` data-zoom-color="${esc(color)}"` : ''}${inRow ? ' aria-hidden="true"' : ` role="button" tabindex="0" aria-label="放大看照片${cap ? `（${esc(cap)}）` : ''}"`}`
    : ''
/** 包一層可以點的框（右下角小放大鏡，看得出可以點） */
const zoomWrap = (inner, attrs, cls = '') => `<span class="zoomable${cls ? ` ${cls}` : ''}"${attrs}>${inner}<span class="zoom-badge" aria-hidden="true">${icon('zoom-in', cls.includes('lg') ? 16 : cls.includes('md') ? 13 : 11)}</span></span>`

/**
 * 兩指縮放（放大檢視共用）：get() 現在幾倍；set(z, cx, cy) 以畫面上 (cx, cy) 那一點為中心縮放。
 * 一指拖曳交給瀏覽器捲動（CSS touch-action: pan-x pan-y），兩指才由這裡處理。
 */
function pinchZoom(scroll, get, set, end) {
  let p = null
  const dist = (t) => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY)
  scroll.addEventListener('touchstart', (e) => {
    if (e.touches.length === 2) p = { d0: dist(e.touches) || 1, z0: get() }
  }, { passive: true })
  scroll.addEventListener('touchmove', (e) => {
    if (!p || e.touches.length !== 2) return
    e.preventDefault()
    const r = scroll.getBoundingClientRect()
    const cx = (e.touches[0].clientX + e.touches[1].clientX) / 2 - r.left
    const cy = (e.touches[0].clientY + e.touches[1].clientY) / 2 - r.top
    set(clampZoom((p.z0 * dist(e.touches)) / p.d0), cx, cy)
  }, { passive: false })
  const stop = (e) => {
    if (p && e.touches.length < 2) {
      p = null
      end?.()
    }
  }
  scroll.addEventListener('touchend', stop)
  scroll.addEventListener('touchcancel', stop)
  // 電腦：觸控板兩指開合（瀏覽器送 ctrl＋滾輪）也能縮放
  scroll.addEventListener('wheel', (e) => {
    if (!e.ctrlKey) return
    e.preventDefault()
    const r = scroll.getBoundingClientRect()
    set(clampZoom(get() * Math.exp(-e.deltaY / 200)), e.clientX - r.left, e.clientY - r.top)
  }, { passive: false })
}

/**
 * 全螢幕看照片：黑底、照片完整顯示；兩指放大、拖曳看細節；點兩下放大／還原。
 * 右上 ×、Esc、點照片旁邊、往下滑都會關；好幾張時左右滑、按 ←／→ 或兩邊的箭頭換張。
 * list：[{ src, cap, places, box, color }]；opts.action：下面多一顆按鈕 { label, icon, cls, run(entry, close) }
 * （照片網址都是縮圖共用的，關掉時不能放掉）
 */
function photoViewer(list, start = 0, opts = {}) {
  list = list.filter((x) => x?.src)
  if (!list.length) return
  let i = Math.max(0, Math.min(start, list.length - 1))
  let z = 1
  let fitted = false // 這一張排好了沒（圖已經在快取裡時 load 也會再來一次，不要把放大到框的倍數蓋掉）
  // 連點兩下縮圖：第二下會落在剛打開的黑色區域 → 打開後一下子之內點黑色的地方不關
  const openedAt = Date.now()
  const opener = document.activeElement
  const el = document.createElement('div')
  el.className = 'viewer pv'
  el.setAttribute('role', 'dialog')
  el.setAttribute('aria-modal', 'true')
  el.setAttribute('aria-label', '放大看照片')
  const multi = list.length > 1
  el.innerHTML = `
    <div class="viewer-bar">
      <span class="zoom-label"></span>
      <button class="btn small secondary" data-z="-1" aria-label="縮小">${icon('minus', 20)}</button>
      <button class="btn small secondary" data-z="1" aria-label="放大">${icon('plus', 20)}</button>
      <button class="viewer-close" aria-label="關閉">${icon('x', 22)}</button>
    </div>
    <div class="viewer-scroll pv-scroll"><div class="pv-pad"><div class="pv-stage"><img alt="" draggable="false"><span class="pv-box" hidden></span></div></div></div>
    ${multi ? `<button class="pv-nav prev" aria-label="上一張">${icon('chevron-left', 26)}</button><button class="pv-nav next" aria-label="下一張">${icon('chevron-right', 26)}</button>` : ''}
    <div class="pv-foot"><div class="pv-cap"></div>${opts.action ? `<button class="btn small ${opts.action.cls || 'secondary'} pv-act">${opts.action.icon ? icon(opts.action.icon, 18) : ''}${esc(opts.action.label)}</button>` : ''}</div>`
  const $ = (s) => el.querySelector(s)
  const scroll = $('.pv-scroll')
  const pad = $('.pv-pad')
  const stage = $('.pv-stage')
  const img = $('img')
  const boxEl = $('.pv-box')
  const label = $('.zoom-label')
  const minus = $('[data-z="-1"]')
  const plus = $('[data-z="1"]')
  // 手機、平板寫「兩指」；電腦寫「點兩下」
  const touch = matchMedia('(hover: none)').matches
  const showLabel = () => {
    label.textContent = `${multi ? `${i + 1} / ${list.length}・` : ''}${z > 1.01 ? `放大 ${z.toFixed(1)} 倍・可以拖曳` : touch ? (multi ? '兩指放大、左右滑換張' : '兩指或點兩下放大') : multi ? '點兩下放大、按左右箭頭換張' : '點兩下放大'}`
    minus.disabled = z <= 1.01
    plus.disabled = z >= ZOOM_MAX - 0.01
  }
  /** 依現在幾倍排版：以 (cx, cy) 那一點為中心（fx、fy 有給就把照片的那個位置放到中間） */
  const layout = (nz, cx, cy, fx, fy) => {
    const sw = scroll.clientWidth
    const sh = scroll.clientHeight
    if (cx == null) cx = sw / 2
    if (cy == null) cy = sh / 2
    const ar = img.naturalWidth && img.naturalHeight ? img.naturalWidth / img.naturalHeight : sw / Math.max(1, sh)
    // 完整顯示（contain）；取整數，避免差零點幾個像素就跑出捲軸
    const bw = Math.floor(Math.min(sw, sh * ar))
    const bh = Math.min(sh, Math.floor(bw / ar))
    const ow = stage.offsetWidth || bw
    const oh = stage.offsetHeight || bh
    if (fx == null) fx = (scroll.scrollLeft + cx - (pad.offsetWidth - ow) / 2) / ow
    if (fy == null) fy = (scroll.scrollTop + cy - (pad.offsetHeight - oh) / 2) / oh
    z = clampZoom(nz)
    scroll.classList.toggle('zoomed', z > 1.01)
    const w = bw * z
    const h = bh * z
    const pw = Math.max(sw, w)
    const ph = Math.max(sh, h)
    stage.style.width = `${w}px`
    stage.style.height = `${h}px`
    pad.style.width = `${pw}px`
    pad.style.height = `${ph}px`
    scroll.scrollLeft = (pw - w) / 2 + fx * w - cx
    scroll.scrollTop = (ph - h) / 2 + fy * h - cy
    showLabel()
  }
  const show = (k) => {
    i = (k + list.length) % list.length
    const it = list[i]
    z = 1
    fitted = false
    img.alt = it.cap || '照片'
    img.src = it.src
    const cap = $('.pv-cap')
    cap.innerHTML = `${(it.places || []).map((p) => locTag(findLocation(p)?.code || p, 'sm')).join('')}${it.cap ? `<span class="pv-cap-text">${esc(it.cap)}</span>` : ''}`
    cap.hidden = !cap.innerHTML
    $('.pv-foot').hidden = cap.hidden && !opts.action
    const box = it.box
    if (box?.length === 4) {
      const [a, b, c, d] = box
      Object.assign(boxEl.style, { top: `${Math.min(a, c) / 10}%`, left: `${Math.min(b, d) / 10}%`, height: `${Math.abs(c - a) / 10}%`, width: `${Math.abs(d - b) / 10}%`, borderColor: it.color || '#0a84ff' })
    }
    boxEl.hidden = !(box?.length === 4)
    if (img.complete && img.naturalWidth) fit()
    else layout(1)
  }
  /** 照片載好：先完整顯示；有框的話直接放大到那一框（框占畫面一半左右） */
  const fit = () => {
    if (fitted) return
    fitted = true
    const it = list[i]
    layout(1, null, null, 0.5, 0.5)
    if (it.box?.length !== 4) return
    const [a, b, c, d] = it.box
    const wf = Math.max(0.02, Math.abs(d - b) / 1000)
    const hf = Math.max(0.02, Math.abs(c - a) / 1000)
    const bw = stage.offsetWidth
    const bh = stage.offsetHeight
    const nz = Math.min((0.5 * scroll.clientWidth) / (wf * bw), (0.5 * scroll.clientHeight) / (hf * bh))
    if (nz > 1.1) layout(nz, null, null, (b + d) / 2000, (a + c) / 2000)
  }
  img.addEventListener('load', fit)
  img.addEventListener('error', () => {
    $('.pv-foot').hidden = false
    $('.pv-cap').hidden = false
    $('.pv-cap').innerHTML = `<span class="pv-cap-text">這張照片打不開（可能還在下載或已經刪掉）</span>`
  })
  const close = () => {
    el.remove()
    window.removeEventListener('keydown', onKey, true)
    window.removeEventListener('resize', onResize)
    document.body.classList.toggle('no-scroll', !!document.querySelector('.viewer'))
    // 焦點回到原本的縮圖；原本在輸入框的不還原（iPhone 會跳出鍵盤）
    if (opener?.isConnected && !opener.matches?.('input, textarea, select, [contenteditable]')) opener.focus?.({ preventScroll: true })
  }
  const onResize = () => layout(z)
  const onKey = (e) => {
    const k = e.key
    // 按住 Enter／空白鍵打開時，重複送出的按鍵不要按到 ×
    if (e.repeat && (k === 'Enter' || k === ' ')) {
      e.preventDefault()
      e.stopImmediatePropagation()
      return
    }
    if (k === 'Escape') close()
    else if (multi && k === 'ArrowLeft' && z <= 1.01) show(i - 1)
    else if (multi && k === 'ArrowRight' && z <= 1.01) show(i + 1)
    else if (k === '+' || k === '=') layout(z + 1)
    else if (k === '-') layout(z - 1)
    else if (k === 'Tab') {
      // 焦點留在放大檢視裡面
      const f = [...el.querySelectorAll('button:not(:disabled)')].filter((b) => b.getClientRects().length) // 手機藏起來的左右箭頭不算
      const at = f.indexOf(document.activeElement)
      const next = e.shiftKey ? (at <= 0 ? f.length - 1 : at - 1) : at < 0 || at >= f.length - 1 ? 0 : at + 1
      f[next]?.focus()
    } else return
    e.preventDefault()
    e.stopImmediatePropagation()
  }
  window.addEventListener('keydown', onKey, true)
  window.addEventListener('resize', onResize)
  $('.viewer-close').onclick = close
  minus.onclick = () => layout(z - 1)
  plus.onclick = () => layout(z + 1)
  if (multi) {
    $('.pv-nav.prev').onclick = () => show(i - 1)
    $('.pv-nav.next').onclick = () => show(i + 1)
  }
  if (opts.action) $('.pv-act').onclick = () => opts.action.run(list[i], close)
  // 點照片旁邊（黑色的地方）＝關；點照片兩下＝放大／還原
  let lastTap = 0
  let dragged = false
  scroll.addEventListener('click', (e) => {
    if (dragged) return (dragged = false)
    if (e.target !== img && !e.target.closest('.pv-box')) return Date.now() - openedAt < 350 ? undefined : close()
    const now = Date.now()
    if (now - lastTap < 320) {
      const r = scroll.getBoundingClientRect()
      layout(z > 1.01 ? 1 : 2.5, e.clientX - r.left, e.clientY - r.top)
      lastTap = 0
    } else lastTap = now
  })
  pinchZoom(scroll, () => z, (nz, cx, cy) => layout(nz, cx, cy), () => z < 1.05 && layout(1))
  // 沒放大時：往下滑＝關，左右滑＝換張（手指跟著動，放開才決定）
  let sw = null
  scroll.addEventListener('touchstart', (e) => {
    sw = e.touches.length === 1 && z <= 1.01 ? { x: e.touches[0].clientX, y: e.touches[0].clientY, dx: 0, dy: 0 } : null
  }, { passive: true })
  scroll.addEventListener('touchmove', (e) => {
    if (!sw || e.touches.length !== 1) return (sw = null), (pad.style.transition = ''), (pad.style.transform = '')
    sw.dx = e.touches[0].clientX - sw.x
    sw.dy = e.touches[0].clientY - sw.y
    const down = sw.dy > 0 && Math.abs(sw.dy) > Math.abs(sw.dx)
    pad.style.transition = 'none' // 手指拖的時候照片要跟著手指，不要慢半拍；放開才用動畫彈回
    pad.style.transform = down ? `translateY(${sw.dy}px)` : multi ? `translateX(${sw.dx}px)` : ''
    el.style.setProperty('--pv-dim', down ? String(Math.max(0.35, 1 - sw.dy / 400)) : '1')
  }, { passive: true })
  scroll.addEventListener('touchend', () => {
    if (!sw) return
    const { dx, dy } = sw
    sw = null
    pad.style.transition = ''
    pad.style.transform = ''
    el.style.setProperty('--pv-dim', '1')
    if (dy > 90 && Math.abs(dy) > Math.abs(dx)) close()
    else if (multi && Math.abs(dx) > 60 && Math.abs(dx) > Math.abs(dy)) show(i + (dx < 0 ? 1 : -1))
  })
  // 電腦：放大後可以用滑鼠按住拖曳
  let drag = null
  stage.addEventListener('pointerdown', (e) => {
    dragged = false
    if (e.pointerType !== 'mouse' || z <= 1.01) return
    drag = { x: e.clientX, y: e.clientY, l: scroll.scrollLeft, t: scroll.scrollTop, moved: false }
    stage.setPointerCapture(e.pointerId)
  })
  stage.addEventListener('pointermove', (e) => {
    if (!drag) return
    if (Math.abs(e.clientX - drag.x) + Math.abs(e.clientY - drag.y) > 4) dragged = true
    scroll.scrollLeft = drag.l - (e.clientX - drag.x)
    scroll.scrollTop = drag.t - (e.clientY - drag.y)
  })
  stage.addEventListener('pointerup', () => (drag = null))
  document.body.appendChild(el)
  document.body.classList.add('no-scroll')
  show(i)
  $('.viewer-close').focus({ preventScroll: true })
  return close
}
/** 結果頁有框的放大檢視（state.viewer）：上面那一行字 */
function reviewZoomLabel() {
  return `放大 ${Number.isInteger(state.zoom) ? state.zoom : state.zoom.toFixed(1)} 倍・可以上下左右滑`
}
/** 結果頁有框的放大檢視：兩指縮放（框是用 % 擺的，跟著一起變大），放開手再重畫按鈕狀態 */
function bindReviewPinch(scroll) {
  const ph = scroll.querySelector('.viewer-photo')
  if (!ph || scroll.dataset.pinch) return
  scroll.dataset.pinch = '1'
  pinchZoom(
    scroll,
    () => state.zoom,
    (nz, cx, cy) => {
      const fx = (scroll.scrollLeft + cx) / Math.max(1, ph.offsetWidth)
      const fy = (scroll.scrollTop + cy) / Math.max(1, ph.offsetHeight)
      state.zoom = nz
      ph.style.width = `${nz * 100}%`
      scroll.scrollLeft = fx * ph.offsetWidth - cx
      scroll.scrollTop = fy * ph.offsetHeight - cy
      const lb = $app.querySelector('.viewer .zoom-label')
      if (lb) lb.textContent = reviewZoomLabel()
      // ctrl＋滾輪縮放不會重畫：＋／－ 能不能按，這裡直接更新
      const out = $app.querySelector('.viewer [data-action="zoom-out"]')
      const inn = $app.querySelector('.viewer [data-action="zoom-in"]')
      if (out) out.disabled = nz <= 1
      if (inn) inn.disabled = nz >= ZOOM_MAX
    },
    () => render(),
  )
}
// 結果頁的放大檢視：Esc 關掉（上面有開視窗時，Esc 先關視窗）
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Escape' || !state.viewer || document.querySelector('.sheet-backdrop, .viewer.pv')) return
  closeReviewViewer()
})
/** 結果頁的放大檢視：打開後焦點移進去（右上 ×）；關掉後回到照片右下的放大鈕 */
async function openReviewViewer() {
  state.viewer = true
  state.zoom = 2
  await render()
  $app.querySelector('.viewer .viewer-close')?.focus({ preventScroll: true })
}
async function closeReviewViewer() {
  state.viewer = false
  await render()
  $app.querySelector('.photo-zoom')?.focus({ preventScroll: true })
}
/** 從縮圖打開：同一組（data-zoom-group）的照片一起放進去，可以左右換張 */
function openZoom(t) {
  const g = t.dataset.zoomGroup
  const els = g ? [...document.querySelectorAll('[data-zoom-src]')].filter((x) => x.dataset.zoomGroup === g) : [t]
  let start = 0
  const list = []
  for (const x of els) {
    if (x === t) start = list.length
    const d = x.dataset
    const places = d.zoomPlaces ? d.zoomPlaces.split('\n') : []
    const box = d.zoomBox ? d.zoomBox.split(',').map(Number) : null
    for (const src of d.zoomSrc.split('\n')) list.push({ src, cap: d.zoomCap || '', places, box, color: d.zoomColor || '' })
  }
  photoViewer(list, start)
}
// 點縮圖：先攔下來（capture），不讓那一列原本的點擊（進詳細頁、選這個產品）跑掉
document.addEventListener('click', (e) => {
  const t = e.target.closest?.('[data-zoom-src]')
  if (!t) return
  e.preventDefault()
  e.stopPropagation()
  openZoom(t)
}, true)
// 鍵盤：Enter 按下就開（按住重複的不算）；空白鍵等放開才開（放開那一下才不會按到剛出現的 ×，Firefox 會）
document.addEventListener('keydown', (e) => {
  if ((e.key !== 'Enter' && e.key !== ' ') || !e.target.matches?.('[data-zoom-src]')) return
  e.preventDefault()
  e.stopPropagation()
  if (e.key === 'Enter' && !e.repeat) openZoom(e.target)
}, true)
document.addEventListener('keyup', (e) => {
  if (e.key !== ' ' || !e.target.matches?.('[data-zoom-src]')) return
  e.preventDefault()
  e.stopPropagation()
  openZoom(e.target)
}, true)

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

/** 改之前先存一份（復原用）：每張照片的框、AI 刪掉的框、數量、手動新增的品項 */
const snapSession = (s) => JSON.stringify({ o: s.photos.map((p) => p.objects), g: s.photos.map((p) => p.aiGone || []), c: s.counts || {}, m: s.manual || [] })
function restoreSession(s, snap) {
  const v = JSON.parse(snap)
  s.photos.forEach((p, i) => {
    p.objects = v.o[i] || []
    p.aiGone = v.g[i] || []
  })
  s.counts = v.c
  s.manual = v.m
}
/** 提示旁邊的「復原」：回到改之前 */
const undoTo = (s, snap) => ({
  label: '復原',
  run: async () => {
    restoreSession(s, snap)
    state.focus = null
    state.focusObj = null
    await save()
    render()
    toast('已復原')
  },
})

/** 修改整個品項（這一列的全部框一起改） */async function editSheet(key) {
  const s = state.session
  const g = groupsOf(s).find((x) => x.key === key)
  if (!g) return
  const sug = await suggestions()
  sheet(
    `<h2 class="sheet-title">${g.manual ? "修改品項" : `修改「${esc(g.label)}」這一種（${g.boxes} 個一起改）`}</h2>
     <p class="sheet-sub">${g.manual ? '手動新增的品項' : `照片裡 ${g.boxes} 個框會一起改；只有其中幾個不一樣，請在照片上點那個框`}</p>
     ${fieldsHtml(g, sug)}
     <div class="row-actions" style="margin-top:16px"><button class="btn" style="flex:1" id="e-save">儲存</button><button class="btn danger" id="e-del">刪掉</button></div>
     ${g.manual || g.boxes < 2 ? '' : `<button class="btn secondary block" id="e-refine" style="margin-top:10px">再比對一次：把這 ${g.boxes} 個分得更細</button><p class="footnote" style="margin:8px 2px 0">把這一列的框切成小圖並排給 AI，一個一個比粗細和接口（要網路，約 10～30 秒）。</p>`}
     ${g.manual ? '' : '<button class="btn secondary block" id="e-sample" style="margin-top:10px">儲存，並存成樣品照</button><p class="footnote" style="margin:8px 2px 0">用照片上第一個框當樣品。這一列混了不同尺寸的話，請先點照片上那一個框、再點一次，從那裡存。</p>'}`,
    (el, close) => {
      el.querySelector('#e-save').onclick = async () => {
        const v = readFields(el)
        if (!v.label) return toast('品名不能空白')
        const snap = snapSession(s)
        if (g.manual) Object.assign(g.manual, v)
        else moveObjects(s, g.refs, v)
        state.focus = null
        await save()
        close()
        render()
        toast(g.manual ? '已修改' : `已修改這一種（${g.boxes} 個）`, undoTo(s, snap))
      }
      el.querySelector('#e-refine')?.addEventListener('click', async (ev) => {
        const btn = ev.currentTarget
        btn.disabled = true
        btn.textContent = '比對中…（約 10～30 秒）'
        try {
          const refs = await sampleRefs()
          const snap = snapSession(s)
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
          toast(groups > 1 ? `分成 ${groups} 種了，請對照照片確認` : 'AI 比對後還是認為是同一種；可以點框個別修改，或存樣品照', groups > 1 ? undoTo(s, snap) : undefined)
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
        const snap = snapSession(s)
        if (g.manual) s.manual = s.manual.filter((m) => m !== g.manual)
        else removeObjects(s, g.refs)
        state.focus = null
        await save()
        close()
        render()
        toast('已刪掉', undoTo(s, snap))
      }
    },
  )
}

/**
 * 點照片上的框：「這一個是哪一種？」按一下就改、改好就關（不會改到別的框）；「N 個要確認」逐一看時才自動跳到下一個。
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
  let formOpen = false
  sheet('<div id="q-body"></div>', (el, close) => {
    const body = el.querySelector('#q-body')
    const finish = () => {
      state.focusObj = null
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
    // 換到下一個框後 0.6 秒內不接受點擊：連點兩下才不會把下一個（本來是對的）也改掉
    let guardUntil = 0
    const guarded = () => Date.now() < guardUntil
    const advance = () => {
      guardUntil = Date.now() + 600
      go(cur + 1)
    }
    // 點到的那一列先打勾 0.25 秒，看得到「有點到」再換畫面；這段時間不接受別的點擊
    const untick = () =>
      body.querySelectorAll('.chosen').forEach((x) => {
        x.classList.remove('chosen')
        x.querySelector('.q-tick')?.remove()
      })
    const tick = async (b) => {
      guardUntil = Date.now() + 1000
      untick()
      b.classList.add('chosen')
      if (b.classList.contains('row')) b.insertAdjacentHTML('beforeend', `<span class="q-tick" aria-hidden="true">${icon('check', 22)}</span>`)
      await sleep(250)
    }
    // 存不進去：打勾拿掉、可以馬上再點（原因 save() 已經跳出來）
    const failed = () => {
      untick()
      guardUntil = 0
    }
    // 開著這個視窗時，另一台更新了這次盤點（畫面已換成新的那份）：舊的框不能再改
    const stale = () => {
      if (state.session === s) return false
      toast('另一台剛更新了這次盤點：已換成最新的，請再點一次框')
      finish()
      return true
    }
    const assign = async (fields) => {
      if (guarded() || stale()) return
      const { pi, o } = order[cur]
      const oi = s.photos[pi].objects.indexOf(o)
      if (oi < 0) return go(cur + 1)
      // 記住改之前的樣子，給「復原」用
      const before = { label: o.label, brand: o.brand, model: o.model, spec: o.spec, edited: o.edited, checked: o.checked }
      const counts = { ...(s.counts || {}) }
      moveObjects(s, [{ pi, oi }], fields)
      try {
        await save()
      } catch {
        Object.assign(o, before)
        s.counts = counts
        return failed()
      }
      render()
      const name = `${fields.label}${fields.spec ? `・${fields.spec}` : ''}`
      const undo = {
        label: '復原',
        run: async () => {
          Object.assign(o, before)
          s.counts = counts
          await save()
          render()
          toast('已復原')
        },
      }
      // 點一個框來改：改好就關掉（不要自動跳到下一個，才不會改到本來就對的）
      if (!doubt) {
        toast(`已把這 1 個改成「${name}」`, undo)
        return finish()
      }
      if (cur + 1 >= order.length) {
        toast(`已改成「${name}」；要確認的都看完了`, undo)
        return finish()
      }
      toast(`已改成「${name}」，換下一個`, undo)
      advance()
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
        <div class="q-head">
          <h2 class="sheet-title">${doubt ? '這一個一樣嗎？' : '這一個是哪一種？'}</h2>
          <span class="q-count"><span class="nb">${doubt ? '要確認的' : ''}第 ${cur + 1} / ${order.length} 個</span>${s.photos.length > 1 ? `・<span class="nb">第 ${pi + 1} 張照片</span>` : ''}</span>
        </div>
        <div class="q-zoom"${zoomAttrs(urlOf(photo), `框起來的這一個・${o.label}${s.photos.length > 1 ? `・第 ${pi + 1} 張照片` : ''}`, { box: o.box, color: mine?.color })}></div>
        ${reason ?`<p class="doubt-reason">為什麼要看：${esc(reason)}</p>` : ''}
        <p class="sheet-sub" style="margin:0 0 10px">目前：${mine ? `<b style="color:${mine.color}">${groups.indexOf(mine) + 1}</b> ${esc(o.label)}${detailOf(o) ? `・${detailHtml(o)}` : ''}` : esc(o.label)}</p>
        ${doubt ? `<button class="btn block" id="q-same" style="margin-bottom:12px">${icon('check', 20)}一樣，就是「${esc(mine ? `${mine.label}${mine.spec ? `・${mine.spec}` : ''}` : o.label)}」</button><p class="sheet-sub" style="margin:0 0 8px">不一樣的話，選它是哪一種：</p>` : ''}
        <div class="group">
          ${groups
            .map(
              (g, gi) => `<button class="row" data-q="${gi}" ${g === mine ? 'aria-current="true"' : ''}>
                <span class="swatch" style="--c:${g.color};pointer-events:none">${gi + 1}</span>
                <span class="grow"><span class="title">${esc(g.label)}</span><br><span class="meta">${detailOf(g) ? detailHtml(g) : '沒有寫尺寸'}・<span class="nb">目前 ${g.boxes} 個</span></span></span>
                ${g === mine ? `<span class="muted now-mark">${icon('check', 16)}目前</span>` : ''}
              </button>`,
            )
            .join('')}
          <button class="row" id="q-new"><span class="add-dot" aria-hidden="true">${icon('plus', 18)}</span><span class="grow"><span class="title">新的一種…</span><br><span class="meta">例如同樣是三通，但尺寸不一樣</span></span></button>
        </div>
        <div id="q-form" style="display:${formOpen ? 'block' : 'none'}"></div>
        <div class="row-actions" style="margin-top:14px">
          <button class="btn secondary" id="q-prev" style="flex:1" ${cur === 0 ? 'disabled' : ''}>${icon('chevron-left', 20)}上一個</button>
          <button class="btn secondary" id="q-next" style="flex:1">${cur + 1 >= order.length ? '完成' : `下一個${icon('chevron-right', 20)}`}</button>
        </div>
        <button class="btn plain block" id="q-more" style="margin-top:8px">改品牌、型號，或存成樣品照…</button>
        <button class="btn danger block" id="q-del" style="margin-top:8px">這不是商品，刪掉這個框</button>`
      const zoom = body.querySelector('.q-zoom')
      zoom.innerHTML = `${boxZoomSvg(photo, o.box, mine?.color, zoom.clientWidth / zoom.clientHeight)}${photo.blob ? `<span class="zoom-badge" aria-hidden="true">${icon('zoom-in', 16)}</span>` : ''}`
      glueTails(body)
      const mineName = mine ? `${mine.label}${mine.spec ? `・${mine.spec}` : ''}` : o.label
      // 確認一樣：記成「看過了」再存；存不進去就還原
      const confirmSame = async () => {
        const was = o.checked
        o.checked = true
        try {
          await save()
          return true
        } catch {
          o.checked = was
          failed()
          return false
        }
      }
      body.querySelector('#q-same')?.addEventListener('click', async (e) => {
        if (guarded() || stale()) return
        await tick(e.currentTarget)
        if (!(await confirmSame())) return
        render()
        if (cur + 1 < order.length) toast('確認了，換下一個')
        advance()
      })
      body.querySelectorAll('[data-q]').forEach((b) =>
        b.addEventListener('click', async () => {
          if (guarded() || stale()) return
          await tick(b)
          const g = groups[Number(b.dataset.q)]
          if (g === mine) {
            // 選了目前這一種＝確認一樣（點一個框來看的：直接關掉）
            if (!(await confirmSame())) return
            if (doubt) {
              if (cur + 1 < order.length) toast('確認了，換下一個')
              return advance()
            }
            toast(`沒有改：本來就是「${mineName}」`)
            return finish()
          }
          guardUntil = 0
          assign({ label: g.label, brand: g.brand, model: g.model, spec: g.spec })
        }),
      )
      body.querySelector('#q-new').onclick = async () => {
        formOpen = true
        const form = body.querySelector('#q-form')
        form.style.display = 'block'
        form.innerHTML = `<p class="sheet-sub" style="margin:12px 2px 0">品名可以一樣，<b>尺寸要填不一樣的</b>（例如「4分 等徑」「5分×3分 異徑」），才會變成新的一種。</p>${fieldsHtml({ label: o.label, brand: o.brand, model: o.model }, await suggestions())}<button class="btn block" id="q-create" style="margin-top:12px">建立並改成這一種</button>`
        glueTails(form)
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
  // sid：從哪次盤點切出來的（用標準答案考 AI 時，那幾次的樣品照不送，不然等於先看答案）
  const sid = state.view === 'review' ? state.session?.id : undefined
  await idb.samples.put({ id: uid(), ...fields, blob, createdAt: Date.now(), ...(sid ? { sid } : {}) })
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
     <div class="sample-new">${zoomWrap(`<img src="${url}" alt="樣品照">`, zoomAttrs(url, '新的樣品照'), 'lg')}</div>
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
       items.map((it) => `<button class="row" data-pick="${esc(it.id)}" data-search="${esc(canon([it.no, it.label, it.brand, it.model, it.spec].join(' ')))}">${itemThumb(it, 36, true)}<span class="grow"><span class="title">${esc(itemTitle(it))}</span><br><span class="meta">${esc([it.no, it.brand, it.model].filter(Boolean).join('・'))}</span></span></button>`).join('') ||
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
/** titleHtml：標題要放儲位黃標籤時用（已經處理過的 HTML）；title 一定要給（讀螢幕的人聽到的） */
function numberSheet({ title, titleHtml, sub, value = '', action = '儲存', quick = [] }, onSave) {
  sheet(
    `<h2 class="sheet-title">${titleHtml ?? esc(title)}</h2><p class="sheet-sub">${sub}</p>
     <input class="field big-num" id="n-val" inputmode="numeric" value="${esc(value)}" aria-label="${esc(title)}">
     ${quick.length ? `<div class="chips" style="margin-top:10px">${quick.map((q) => `<button class="chip" data-n="${q.n}">${esc(q.label)}</button>`).join('')}</div>` : ''}
     <div class="row-actions" style="margin-top:14px"><button class="btn secondary" id="n-cancel" style="flex:1">取消</button><button class="btn" id="n-save" style="flex:1">${esc(action)}</button></div>`,
    (el, close) => {
      el.querySelector('#n-cancel').onclick = close
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

/** 複盤後選原因（選一個＋可以補一句說明）；會寫進試算表「盤差報告」的原因欄 */
function reasonSheet(it, sid) {
  const rc = it.recounts[sid]
  let pick = rc.reason || ''
  sheet(
    `<h2 class="sheet-title">差異的原因</h2><p class="sheet-sub">${esc(itemTitle(it))}・${placeHtml(rc.place)}。選最有可能的一個；不確定就選「其他」，寫一句說明。</p>
     <div class="chips" id="r-chips">${REASONS.map((r) => `<button class="chip" data-r="${esc(r)}" aria-pressed="${r === pick}">${esc(r)}</button>`).join('')}</div>
     <input class="field" id="r-note" placeholder="補充說明（選填），例如：上週借給客人兩個" value="${esc(rc.note || '')}" style="margin-top:12px">
     <button class="btn block" id="r-save" style="margin-top:14px">儲存原因</button>`,
    (el, close) => {
      el.querySelectorAll('[data-r]').forEach((b) =>
        b.addEventListener('click', () => {
          pick = b.dataset.r
          el.querySelectorAll('[data-r]').forEach((x) => x.setAttribute('aria-pressed', String(x === b)))
        }),
      )
      el.querySelector('#r-save').onclick = async () => {
        if (!pick) return toast('先選一個原因')
        Object.assign(rc, { reason: pick, note: el.querySelector('#r-note').value.trim(), reasonBy: ls.get(LS.memberName) || '', v: Date.now() })
        close()
        await putItem(it)
        toast(canManage() ? '原因已存：可以按「調整帳面」或「不調整」' : '原因已存：等擁有者或管理員決定要不要調整帳面')
        render()
      }
    },
  )
}

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
       ${canManage() ? `<button class="row" id="m-erp"><span class="grow"><span class="title">匯入正航產品表</span><br><span class="meta">產品存量明細表（CSV 或貼上）：產品總表、帳面數</span></span>${chev}</button>` : ''}
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
      el.querySelector('#m-erp')?.addEventListener('click', () => {
        close()
        erpImportSheet()
      })
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
        downloadBlob(await backupBlob(), `聖佳庫存備份_${ymd(Date.now())}.json`)
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
  [`叫貨清單 ${ymd(Date.now())}`, ...list.map((it) => `・${itemTitle(it)}${it.brand || it.model ? `（${[it.brand, it.model].filter(Boolean).join(' ')}）` : ''}${blindMe() ? '' : `：剩 ${orderQty(it)} 個`}（設定剩 ${it.safety} 個以下要叫貨）`)].join('\n')

/** 品項庫的表格（Excel、Google 試算表共用） */
const ITEM_HEAD = ['料號', '品名', '品牌', '型號', '尺寸／規格', '實盤', '帳面', '差異', '剩幾個要叫貨（安全庫存）', '狀態', '在哪裡（位置 數量）', '最近盤點', '盤點人', '複盤', '差異原因']
/** 每一格「誰盤的」：從那次盤點找盤點人 */
async function whoOfStock() {
  const byId = new Map((await db.all()).map((s) => [s.id, s]))
  return (st) => (byId.get(st.sid) ? byName(byId.get(st.sid)) : '')
}
const timeText = (ms) => (ms ? `${ymd(ms)} ${hm(ms)}` : '')
/** time：時間怎麼寫（Excel 用文字；Google 試算表送 {$t} 讓那邊轉成真的日期）；hide：盲盤的人下載時，帳面、差異留空 */
function itemRows(items, whoOf = () => '', time = timeText, hide = false) {
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  return [...items]
    .sort((a, b) => cmp(a.label, b.label) || cmp(a.spec, b.spec) || cmp(a.no, b.no))
    .map((it) => {
      const d = hide ? null : diffOf(it)
      const rc = latestRecount(it)
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
        hide ? '' : (it.book ?? ''),
        d ?? '',
        it.safety ?? '',
        [it.status === 'new' ? '新的（待確認）' : '', needsOrder(it) ? '該叫貨' : '', d > 0 ? '盤盈' : d < 0 ? '盤虧' : ''].filter(Boolean).join('、'),
        liveStock(it)
          .map(([, st]) => `${st.place || '沒填位置'} ${st.count}`)
          .join('、'),
        last ? time(last) : '',
        latest ? whoOf(latest) : '',
        rc ? RC_TEXT[rc.status] : '',
        rc?.reason ? `${rc.reason}${rc.note ? `：${rc.note}` : ''}` : '',
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
  { name: '品項庫', rows: [ITEM_HEAD, ...itemRows(items, whoOf, timeText, blindMe())] },
  { name: '儲位庫存', rows: [STOCK_HEAD, ...stockRows(items, whoOf)] },
  { name: '叫貨清單', rows: [['料號', '品名', '品牌', '型號', '尺寸／規格', '現在剩（實盤、帳面取少的）', '剩幾個要叫貨（安全庫存）'], ...items.filter(needsOrder).map((it) => [it.no, it.label, it.brand, it.model, it.spec, blindMe() ? '' : orderQty(it), it.safety])] },
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
  if (data?.app !== '拍照盤點' || data.kind !== 'backup') return toast('這不是聖佳智慧庫存的備份檔')
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
    .map((l) => `<div class="label"><div class="label-code">${esc(l.code)}</div>${l.name ? `<div class="label-name">${esc(l.name)}</div>` : ''}<div class="label-foot">聖佳智慧庫存・儲位</div></div>`)
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
/** 雲端的程式碼是新版（照片分開存，同步只拿資料）；舊版要請擁有者重新貼程式碼、部署新版本 */
const serverV2 = () => Number(ls.get(LS.serverVer, '0')) >= 2
const serverOld = () => syncReady() && !!ls.get(LS.lastSync) && !serverV2()
let syncTimer = 0
let syncing = null
let repairing = null
/** 把雲端送來的照片放進這台的那次盤點（不改修改時間：這台改過、還沒傳的，照常上傳）；回傳補了幾張 */
async function putPhotos(id, got) {
  const cur = await db.get(id)
  if (!cur) return 0
  let n = 0
  const fill = (p) => {
    const b64 = !p.blob && got[p.id]
    if (!b64) return false
    p.blob = b64ToBlob(b64)
    delete p.lost
    delete p.pending
    p.up = true
    return true
  }
  for (const p of cur.photos) if (fill(p)) n++
  if (!n) return 0
  await db.putRaw(cur)
  // 正在看這次盤點：畫面上的那份也補上
  if (state.session?.id === id) {
    state.session.photos.forEach(fill)
    if (['review', 'capture'].includes(state.view)) render()
  }
  return n
}
/** 雲端那一筆有這台不見的照片（舊格式，照片跟資料綁一起）：補回來；回傳補了幾張 */
async function fillLostPhotos(rec) {
  if (rec.del || !rec.d?.photos || !rec.k.startsWith('session:')) return 0
  const got = {}
  for (const p of rec.d.photos) if (p.b64) got[p.id] = p.b64
  return Object.keys(got).length ? putPhotos(rec.k.slice('session:'.length), got) : 0
}
/**
 * 抓一次盤點的照片（還沒下載的、這台弄丟的）：同步平常只拿資料，照片打開那次盤點才抓這一次的。
 * 一台同時只抓一次；回傳抓到幾張。
 */
const fetching = new Map()
function fetchPhotos(s, { quiet = true } = {}) {
  const want = s.photos.filter((p) => !p.blob && (p.pending || p.lost)).map((p) => p.id)
  if (!want.length || !syncReady()) return Promise.resolve(0)
  if (fetching.has(s.id)) return fetching.get(s.id)
  const job = (async () => {
    const got = {}
    let keys = want.map((pid) => photoKey(s.id, pid))
    // 一次最多約 8 MB：沒拿完的再要
    for (let round = 0; round < 20 && keys.length; round++) {
      const r = await postSync({ action: 'photos', keys })
      // postSync 遇到雲端說不行會直接丟錯誤；這裡只剩「舊版程式碼不認得 photos」的情況
      if (!Array.isArray(r.records)) throw new Error('雲端的程式碼是舊版，還不會傳照片：請擁有者更新程式碼（設定 → 複製試算表程式碼）')
      const back = new Set()
      for (const rec of r.records || []) {
        back.add(rec.k)
        if (rec.d?.b64) got[rec.k.slice(rec.k.lastIndexOf(':') + 1)] = rec.d.b64
      }
      keys = keys.filter((k) => !back.has(k))
      if (!r.more) break
    }
    const n = await putPhotos(s.id, got)
    if (!quiet) toast(n ? `拿回 ${n} 張照片` : '雲端也沒有這幾張照片：請在拍照的那台手機打開 App，按「立即同步」')
    return n
  })().finally(() => fetching.delete(s.id))
  fetching.set(s.id, job)
  return job
}
/** 同步完，趁沒事慢慢把還沒下載的照片抓下來（新的先），首頁縮圖才會慢慢出現；一次一筆、不影響使用 */
let prefetching = false
async function prefetchPhotos() {
  if (prefetching || !syncReady()) return
  prefetching = true
  try {
    for (let i = 0; i < 200; i++) {
      if (document.visibilityState !== 'visible' || state.view === 'analyzing' || state.busy) break
      const s = (await db.all()).find((x) => x.photos.some((p) => p.pending))
      if (!s) break
      try {
        await fetchPhotos(s)
      } catch {
        break // 連不上就先不要一直試，下次同步再說
      }
      if (state.view === 'home') render()
      await sleep(300)
    }
  } finally {
    prefetching = false
  }
}
/**
 * 從雲端重新下載全部：平常的同步只拿「上次之後的新資料」、而且跳過這台自己傳的；
 * 這台的資料如果不見了（iPhone 的問題），平常的同步補不回來 → 從頭再拿一次（dev 用別的名字，連這台傳的也拿）。
 * 只補「這台少了的、雲端比較新的、這台不見的照片」；這台比較新的不會被蓋掉。quiet：自動跑的，沒補到就不吵。
 */
function fullPull({ quiet = false } = {}) {
  if (repairing) return repairing
  repairing = (async () => {
    if (!syncReady()) {
      if (!quiet) toast('這台沒有開啟多台同步：雲端沒有資料可以下載')
      return { records: 0, photos: 0 }
    }
    if (!quiet) {
      toast('正在從雲端重新下載…（照片多的話要等一下）')
      render()
    }
    let records = 0
    let photos = 0
    let cursor = 0
    for (let round = 0; round < 200; round++) {
      const r = await postSync({ action: 'pull', since: cursor, dev: 'full-pull' })
      for (const rec of r.records) {
        // 一筆壞掉的不能卡住全部
        try {
          if (await applyRemote(rec)) records++
          else photos += await fillLostPhotos(rec)
        } catch (e) {
          console.error('fullPull', rec?.k, e)
        }
      }
      cursor = r.next
      if (!r.more) break
    }
    itemsCache = null
    if (photos) scheduleSync(2000)
    if (records || photos) toast(`從雲端補回 ${[records && `${records} 筆資料`, photos && `${photos} 張照片`].filter(Boolean).join('、')}`)
    else if (!quiet) toast((await db.all()).some((s) => s.photos.some((p) => p.lost)) ? '雲端也沒有不見的那幾張照片：請在拍照的那台手機打開 App，按「立即同步」' : '這台跟雲端一樣，沒有少')
    return { records, photos }
  })().finally(() => {
    repairing = null
    render()
  })
  return repairing
}
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
/** 雲端（Apps Script）回的錯誤：data＝雲端的回覆（看得出是檢視者、被移除、次數到上限，還是網路問題） */
class SyncError extends Error {
  constructor(message, data = {}) {
    super(message)
    this.data = data
  }
}
/** opt.signal：等太久或使用者取消時中斷（只是這邊不等了；雲端那邊收到的照樣會跑完） */
async function postSync(body, opt = {}) {
  let res
  try {
    // dev：雲端用來跳過「這台自己傳的」；拿回照片時改用別的名字，連自己傳的也要（權限看 key，不看 dev）
    res = await fetch(ls.get(LS.sheet), { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ ...body, key: ls.get(LS.syncKey), dev: body.dev || deviceId() }), signal: opt.signal })
  } catch (e) {
    if (opt.signal?.aborted) throw new SyncError('已取消', { aborted: true })
    throw new SyncError('連不到 Google（沒有網路？）', { network: true })
  }
  const text = await res.text().catch(() => '')
  let data
  try {
    data = JSON.parse(text)
  } catch {
    data = null
  }
  if (!data || typeof data !== 'object') {
    // 不是 JSON：Apps Script 出錯時回的是網頁（例如 Google 的每日用量上限、執行太久），不是程式碼舊
    const login = /accounts\.google\.com|ServiceLogin/i.test(res.url + text.slice(0, 2000))
    const msg = login
      ? 'Google 試算表要登入才能用：部署時「誰可以存取」要選「所有人」，再部署一次新版本'
      : res.status === 404
        ? '找不到這個 Google 試算表網址：請擁有者重新複製「網頁應用程式網址」'
        : 'Google 試算表那邊暫時出錯（可能是 Google 每天的用量上限或執行太久），等一下再試'
    throw new SyncError(msg, { html: true, httpStatus: res.status, login })
  }
  // 擁有者把資料搬到新的 Google 帳號：自動改用新網址（連結碼不用重貼）
  if (data.moved && !body.followed) {
    ls.set(LS.sheet, data.moved)
    toast('資料搬到新的位置了，已自動跟過去')
    return postSync({ ...body, followed: true }, opt)
  }
  // 被移除權限：這台的公司資料自動清掉
  if (data.revoked === true && !data.ok) {
    await wipeLocal('這台已經被移除權限，App 裡的資料已清除')
    throw new SyncError('這台已經被移除權限，App 裡的資料已清除', data)
  }
  // 雲端程式碼的版本（2＝會把照片分開存）：決定上傳用哪種格式、要不要提醒更新程式碼
  if (data.ver) ls.set(LS.serverVer, String(data.ver))
  // 擁有者有沒有把 AI 金鑰放在雲端共用（新版雲端程式碼才會回這個）
  if ('ai' in data) {
    const was = ls.get(LS.aiShared) === '1'
    ls.set(LS.aiShared, data.ai ? '1' : '')
    if (was !== !!data.ai) setTimeout(render, 0)
  }
  // 權限隨時可能被管理員改：每次回覆都更新
  if (data.me?.role) {
    const changed = data.me.role !== ls.get(LS.memberRole)
    ls.set(LS.memberRole, data.me.role)
    ls.set(LS.memberName, data.me.name || '')
    ls.set(LS.memberId, data.me.id || '')
    // 不再是擁有者／管理員：交叉比對的結果（有產品名稱、只有管理的人看）和開關一起清掉
    if (!canManage()) {
      ls.set(LS.pickCrossLast, '')
      ls.set(LS.pickCross, '')
    }
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
  if (!data.ok) throw new SyncError(data.error || '同步失敗', data)
  return data
}
/** 清除這台的資料（退出同步、被移除權限時） */
async function wipeLocal(msg) {
  clearTimeout(syncTimer)
  for (const k of [LS.aiShared, LS.syncKey, LS.sheet, LS.pulled, LS.lastSync, LS.deleted, LS.settingsAt, LS.settingsSyncT, LS.locations, LS.key, LS.catalog, LS.memberName, LS.memberRole, LS.memberId, LS.roster, LS.counterId]) ls.set(k, '')
  // 直接清本機（不留刪除紀錄，才不會把雲端的資料也刪掉）
  await idb.sessions.clear()
  await idb.items.clear()
  await idb.samples.clear()
  await idb.picks.clear().catch(() => {})
  // 交叉比對的結果裡有產品名稱（公司資料）：一起清掉
  ls.set(LS.pickCrossLast, '')
  itemsCache = null
  picksCache = null
  state.pick = null
  state.pickEdit = null
  state.session = null
  toast(msg)
  go('home')
}
// pick：點貨紀錄（4.7）。雲端的程式碼不用改：push／pull 本來就收任何鍵（只有 erp: 限擁有者、管理員）；
// 舊版 App 收到 pick: 會略過（不認得的種類 applyRemote 回 false），不會出錯
const SYNC_STORES = { session: idb.sessions, item: idb.items, sample: idb.samples, erp: idb.erp, pick: idb.picks }
const tOf = (v) => v.updatedAt || v.createdAt || 0
/**
 * 盤點紀錄（第 2 版格式，v:2）：只傳資料，照片另外一筆一張（photo:盤點ID:照片ID），下載時先拿資料、打開那次盤點才抓照片。
 * 舊格式（照片 b64 跟資料綁在一起）下載時還認得。
 */
const photoKey = (sid, pid) => `photo:${sid}:${pid}`
async function encodeRecord(kind, v, { full = false } = {}) {
  // full：雲端的程式碼還是舊版（還不會收單獨的照片）→ 照片照舊跟資料綁一起傳
  if (kind === 'session' && full) return { ...v, photos: await Promise.all(v.photos.map(async ({ lost, pending, up, ...p }) => ({ ...p, blob: undefined, b64: p.blob ? await blobToBase64(p.blob) : '' }))) }
  if (kind === 'session') return { ...v, v: 2, photos: v.photos.map(({ blob, lost, pending, up, ...p }) => p) }
  if (kind === 'item') return { ...v, photo: undefined, photoB64: v.photo ? await blobToBase64(v.photo) : '' }
  if (kind === 'erp') return { ...v, blob: undefined } // 正航產品表：純資料
  if (kind === 'pick') {
    const { _syncT, ...rest } = v // 點貨紀錄：純資料（沒有客戶名稱、價格）
    return rest
  }
  return { ...v, blob: undefined, b64: await blobToBase64(v.blob) }
}
function decodeRecord(kind, d) {
  if (kind === 'pick') return normPick(d, d?.id)
  if (kind === 'session')
    return {
      ...d,
      photos: (d.photos || []).map(({ b64, ...p }) => (b64 ? { ...p, blob: b64ToBlob(b64) } : d.v >= 2 ? { ...p, blob: undefined, pending: true } : { ...p, blob: undefined, lost: true })),
    }
  if (kind === 'item') {
    const { photoB64, ...rest } = d
    return { ...rest, photo: photoB64 ? b64ToBlob(photoB64) : undefined }
  }
  if (kind === 'erp') {
    const { blob, b64, ...rest } = d
    return rest
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
  // 正航料號：比較晚接上（或取消）的為準；都沒記時間（舊資料）就留有值的那份——
  // 不然一台接上正航、另一台同時改了別的欄位，同步一合併，正航料號就不見了
  const erpSrc = (local.erpAt || 0) === (remote.erpAt || 0) ? null : (local.erpAt || 0) > (remote.erpAt || 0) ? local : remote
  const erpNo = erpSrc ? erpSrc.erpNo || '' : base.erpNo || local.erpNo || remote.erpNo || ''
  return {
    ...base,
    stock,
    moves,
    book,
    erpNo,
    erpAt: Math.max(local.erpAt || 0, remote.erpAt || 0) || undefined,
    unit: base.unit || local.unit || remote.unit,
    aliases: [...new Set([...(local.aliases || []), ...(remote.aliases || [])])],
    equiv: [...new Set([...(local.equiv || []), ...(remote.equiv || [])])],
    recounts: mergeRecounts(local.recounts, remote.recounts),
    photo: base.photo || local.photo || remote.photo,
  }
}
/** 比較兩個版本內容是否一樣（不看照片、時間） */
const itemSig = (it) =>
  JSON.stringify([
    it.no,
    it.erpNo || '',
    it.unit || '',
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
    Object.entries(it.recounts || {})
      .map(([k, rc]) => `${k}|${rc.v}`)
      .sort(),
  ])
/** 套用一筆別台的資料；回傳本機有沒有變 */
async function applyRemote(rec) {
  if (rec.k === 'settings') {
    if (!rec.d || rec.t <= Number(ls.get(LS.settingsAt, '0'))) return false
    ls.set(LS.locations, JSON.stringify(rec.d.locations || []))
    if (rec.d.catalog != null) ls.set(LS.catalog, rec.d.catalog)
    if (rec.d.blind != null) ls.set(LS.blind, rec.d.blind)
    if (rec.d.recount != null) ls.set(LS.recount, rec.d.recount)
    ls.set(LS.settingsAt, String(rec.t))
    ls.set(LS.settingsSyncT, String(rec.t))
    return true
  }
  const i = rec.k.indexOf(':')
  const kind = rec.k.slice(0, i)
  const id = rec.k.slice(i + 1)
  const store = SYNC_STORES[kind]
  if (!store) return false // photo:… 之類的不在這裡處理（照片是打開那次盤點才抓）
  const cur = await store.get(id)
  if (rec.del) {
    if (!cur || tOf(cur) > rec.t) return false
    await store.del(id)
    if (kind === 'session' && state.session?.id === id && ['review', 'capture'].includes(state.view)) {
      toast('這次盤點在另一台被刪掉了')
      go('home')
    }
    if (kind === 'pick') {
      picksCache = null
      if (state.pick?.id === id) {
        // 正在看或正在改這張：標成已刪（AI 讀完、存檔都不會再把它存回去）
        state.pick._gone = true
        if (['pick-edit', 'pick-result'].includes(state.view)) {
          toast('這筆點貨紀錄在另一台被刪掉了')
          state.pick = null
          state.pickEdit = null
          go('pick')
        }
      }
    }
    return true
  }
  if (!rec.d) return false
  if (kind === 'pick') return applyRemotePick(rec, cur, id)
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
  // 盤點：雲端那份只有資料，照片這台已經有的就留著（不用再抓一次）
  if (kind === 'session' && cur) {
    for (const p of v.photos) {
      const mine = cur.photos.find((x) => x.id === p.id)
      if (!p.blob && mine?.blob) Object.assign(p, { blob: mine.blob, pending: undefined, lost: undefined, up: mine.up })
    }
  }
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
  // 別台匯入了新的正航產品表：下次用的時候重新讀（以前要重新打開 App 才看得到）
  if (kind === 'erp') erpCache = undefined
  return true
}
/** 同步一次（同時只跑一個）：先上傳、再下載 */
/** 同步進度：首頁左上的雲朵、設定頁都看得到（不用整頁重畫） */
function syncProgress(msg) {
  state.syncMsg = msg
  document.querySelectorAll('.sync-text').forEach((el) => (el.textContent = msg))
}
async function syncNow(report = () => {}) {
  if (!syncReady()) throw new Error('還沒開啟多台同步')
  if (syncing) return syncing
  const onProgress = (m) => {
    syncProgress(m)
    report(m)
  }
  syncing = (async () => {
    state.syncState = 'syncing'
    onProgress('連線中…')
    // 1. 上傳：還沒同步過的（檢視者只下載）
    const jobs = []
    const viewer = !canEdit()
    for (const s of await db.all()) if (tOf(s) !== s._syncT && s.photos.some((p) => p.status === 'done')) jobs.push(['session', s])
    for (const it of await itemsAll(true)) if (tOf(it) !== it._syncT) jobs.push(['item', it])
    for (const sm of await idb.samples.all().catch(() => [])) if (tOf(sm) !== sm._syncT) jobs.push(['sample', sm])
    for (const e of await idb.erp.all().catch(() => [])) if (tOf(e) !== e._syncT) jobs.push(['erp', e])
    // 點貨紀錄：按過「完成」的才傳（點到一半的只在這台）；有改過才讀（不用每 40 秒把全部點貨紀錄讀一遍）
    const scanPicks = picksDirty
    picksDirty = false
    if (scanPicks) for (const pk of await idb.picks.all().catch(() => [])) if (pk.doneAt && tOf(pk) !== pk._syncT) jobs.push(['pick', pk])
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
      // 雲端沒收（只有擁有者、管理員能改的）：告訴使用者；被擋下的刪除，等一下從雲端把紀錄拿回來，跟大家一致
      if (r.ignored?.length) {
        if (r.ignoredMsg) toast(r.ignoredMsg)
        if (r.ignored.some((k) => /^(session|pick):/.test(String(k)))) ls.set(LS.needRepair, '1')
      }
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
    const full = !serverV2() // 雲端程式碼還是舊版：照片跟資料綁一起傳（舊方式）
    for (const [kind, v] of jobs) {
      const t = tOf(v)
      // 某一筆的照片讀不出來：跳過那一筆，其他照樣同步（以前會整個同步失敗）
      let d
      try {
        d = await encodeRecord(kind, v, { full })
      } catch (e) {
        console.error('encode', kind, v.id, e)
        continue
      }
      // 盤點的照片：還沒傳過的，一張一筆先傳（照片不會改，傳過一次就好）
      if (kind === 'session' && !full) {
        for (const p of v.photos) {
          if (p.up || !p.blob) continue
          let b64
          try {
            b64 = await blobToBase64(p.blob)
          } catch {
            continue
          }
          const pid = p.id
          await add({ k: photoKey(v.id, pid), t, d: { b64, type: p.blob.type || 'image/jpeg' } }, async () => {
            await db.markPhotoUp(v.id, [pid])
            if (state.session?.id === v.id) for (const x of state.session.photos) if (x.id === pid) x.up = true
          })
        }
      }
      await add({ k: `${kind}:${v.id}`, t, d }, async () => {
        await SYNC_STORES[kind].markSynced(v.id, t)
        if (kind === 'item') {
          const c = itemsCache?.find((x) => x.id === v.id)
          if (c && tOf(c) === t) c._syncT = t
        }
        if (kind === 'session' && state.session?.id === v.id && tOf(state.session) === t) state.session._syncT = t
      })
    }
    if (settingsDirty && !viewer) await add({ k: 'settings', t: settingsAt, d: { locations: locations(), catalog: ls.get(LS.catalog), blind: ls.get(LS.blind, '1'), recount: ls.get(LS.recount, '1') } }, async () => ls.set(LS.settingsSyncT, String(settingsAt)))
    for (const d of deleted)
      await add({ k: d.k, t: d.t, del: true }, async () => {
        ls.set(LS.deleted, JSON.stringify(readJson(LS.deleted, []).filter((x) => !(x.k === d.k && x.t === d.t))))
      })
    await flush()
    // 2. 下載：別台送上去、比上次新的
    // 點貨紀錄（4.7）：更新前的舊版 App 下載到別人的點貨紀錄會略過、之後也不會再下載 → 更新後做一次：
    // 把「下載到哪裡」退回 4.7 上線前（雲端的順序號就是收到的毫秒時間），從那天之後的再拿一次；
    // 這台已經有的、比較新的不會被蓋掉，也不會跳提示
    if (ls.get(LS.pickPullDone) !== '1') {
      ls.set(LS.pulled, String(Math.min(Number(ls.get(LS.pulled, '0')) || 0, PICK_LAUNCH)))
      ls.set(LS.pickPullDone, '1')
    }
    let cursor = Number(ls.get(LS.pulled, '0'))
    const fromStart = cursor === 0 // 剛加入、第一次：本來就是從頭拿全部
    let got = 0
    let changed = false
    for (let round = 0; round < 200; round++) {
      onProgress(got ? `下載了 ${got} 筆，還有…` : fromStart ? `下載中（第一次會比較久${serverV2() ? '' : '，照片都要抓下來'}）` : '下載中…')
      const r = await postSync({ action: 'pull', since: cursor })
      for (const rec of r.records) {
        // 一筆壞掉的資料不能卡住後面全部：跳過那一筆、記在主控台
        try {
          if (await applyRemote(rec)) changed = true
        } catch (e) {
          console.error('applyRemote', rec?.k, e)
        }
        got++
        onProgress(`下載 ${got} 筆…`)
      }
      cursor = r.next
      ls.set(LS.pulled, String(cursor))
      if (!r.more) break
    }
    ls.set(LS.lastSync, String(Date.now()))
    // 剛從頭下載過一次，就不用再「重新下載全部」（不然剛加入的人會下載兩次，等很久）
    if (fromStart) ls.set(LS.fullPullDone, '1')
    // 自動從雲端重新下載一次全部：這台有照片不見了（轉新存法時發現），或剛更新到 4.0.7（之前同步卡住，這台可能少了資料）
    if (ls.get(LS.needRepair) === '1' || ls.get(LS.fullPullDone) !== '1') {
      onProgress('檢查有沒有少的紀錄…')
      ls.set(LS.needRepair, '')
      ls.set(LS.fullPullDone, '1')
      const r = await fullPull({ quiet: true }).catch(() => null)
      if (r?.records || r?.photos) changed = true
    }
    state.syncState = ''
    if (changed) {
      itemsCache = null
      picksCache = null
      // 點貨正在點的那一頁（pick-edit）不重畫：不要打斷輸入
      if (['home', 'items', 'item', 'report', 'locations', 'review', 'pick', 'pick-result'].includes(state.view)) render()
    } else if (state.view === 'home') render()
    // 正在看的那次盤點先抓照片，其他的在背景慢慢抓
    if (state.session?.photos.some((p) => p.pending)) fetchPhotos(state.session).catch(() => {})
    setTimeout(() => prefetchPhotos().catch(() => {}), 1500)
    return { pushed: sent, pulled: got, changed }
  })()
    .catch((e) => {
      state.syncState = 'error'
      state.syncError = e.message
      picksDirty = true // 這次沒傳成功：下次再檢查一次點貨紀錄
      if (state.view === 'home') render()
      throw e
    })
    .finally(() => (syncing = null))
  return syncing
}
/** 連結碼＝網址＋同步密碼（給另一台貼上） */
const linkCode = () => `${ls.get(LS.sheet)}#k=${ls.get(LS.syncKey)}`
/** 連結碼兩種寫法都認：舊的（試算表網址#k=密碼）、一鍵加入連結（App 網址#join=試算表代號~密碼） */
function parseLinkCode(text) {
  const s = String(text).trim()
  const m = /^(https:\/\/script\.google\.com\/macros\/s\/[^#\s]+\/exec)#k=([\w-]{12,})$/.exec(s)
  if (m) return { url: m[1], key: m[2] }
  const j = /#join=([\w-]{20,})~([\w-]{12,})$/.exec(s)
  return j ? { url: `https://script.google.com/macros/s/${j[1]}/exec`, key: j[2] } : null
}
/**
 * 一鍵加入的連結：同事在 LINE 點一下就打開 App、按「加入」就好（不用複製貼上）。
 * openExternalBrowser=1：LINE 會改用手機的 Safari／Chrome 打開（在 LINE 內建瀏覽器加入，之後從桌面打開會變成沒加入）。
 * 密碼放在 # 後面：不會送到 GitHub 的伺服器。
 */
function joinLinkOf(code, appUrl) {
  const c = parseLinkCode(code)
  const id = c && /\/macros\/s\/([^/]+)\/exec$/.exec(c.url)?.[1]
  return id ? `${appUrl}?openExternalBrowser=1#join=${id}~${c.key}` : code
}
/** 加入共用：連到雲端、從頭下載一次 */
async function joinWith(code) {
  const before = { url: ls.get(LS.sheet), key: ls.get(LS.syncKey) }
  ls.set(LS.sheet, code.url)
  ls.set(LS.syncKey, code.key)
  try {
    await postSync({ action: 'hello' })
  } catch (err) {
    ls.set(LS.sheet, before.url)
    ls.set(LS.syncKey, before.key)
    throw err
  }
  ls.set(LS.pulled, '0') // 新連結：從頭下載一次
  render()
  syncNow((m) => toast(m))
    .then((r) => toast(`同步完成：下載 ${r.pulled} 筆、上傳 ${r.pushed} 筆`))
    .catch((err) => toast(`同步沒成功：${err.message}`))
    .finally(() => render())
}
const inLineApp = () => /\bLine\//i.test(navigator.userAgent)
const standalone = () => window.matchMedia?.('(display-mode: standalone)').matches || navigator.standalone === true
/** 點了邀請連結打開 App：跳出「加入」（不用到設定裡貼） */
function joinSheet(code) {
  const clean = () => history.replaceState(null, '', location.pathname) // 網址列不要留著密碼
  if (ls.get(LS.syncKey) === code.key && ls.get(LS.sheet) === code.url) {
    clean()
    return toast('這台已經加入了，不用再加一次')
  }
  if (inLineApp())
    return sheet(
      `<h2 class="sheet-title">請用 ${/iPhone|iPad/.test(navigator.userAgent) ? 'Safari' : 'Chrome'} 打開</h2>
       <p class="sheet-sub">現在是在 LINE 裡面打開的。在這裡加入的話，之後從手機桌面打開會變成沒加入。</p>
       <ol class="steps-list"><li>按右上角的「⋯」（或右下角的 <b>⤴︎</b>）。</li><li>選「用預設瀏覽器開啟」（或「在 Safari 中打開」）。</li><li>打開後按「加入」就好。</li></ol>`,
    )
  if (ls.get(LS.syncKey)) {
    clean()
    return sheet(`<h2 class="sheet-title">這台已經加入別的共用</h2><p class="sheet-sub">要換的話，先到「設定 → 這台退出並清除資料」，再點一次邀請連結。</p>`)
  }
  clean()
  sheet(
    `<img class="join-logo" src="logo.svg" alt="聖佳 LOGO"><h2 class="sheet-title" style="text-align:center">加入聖佳智慧庫存</h2>
     <p class="sheet-sub">你被邀請一起用「聖佳智慧庫存」。加入後，大家盤點的資料會自動同步到這支手機。</p>
     <button class="btn block" id="j-go" style="margin-top:8px">加入</button>
     <p class="footnote" style="margin-top:10px">這個邀請只給你一個人用，請不要轉傳。</p>`,
    (el, close) => {
      el.querySelector('#j-go').onclick = async (ev) => {
        ev.currentTarget.disabled = true
        ev.currentTarget.textContent = '加入中…'
        try {
          await joinWith(code)
        } catch (err) {
          close()
          return toast(err.message)
        }
        close()
        const ios = /iPhone|iPad/.test(navigator.userAgent)
        sheet(
          `<h2 class="sheet-title">加入好了</h2>
           <p class="sheet-sub">${esc(ls.get(LS.memberName) || '')}${ls.get(LS.memberName) ? '，' : ''}資料正在下載，等一下就會出現。</p>
           ${
             standalone()
               ? ''
               : `<p class="section-title" style="margin-top:14px">把 App 放到手機桌面，之後比較好找</p>
                  <ol class="steps-list">${ios ? '<li>按畫面下方的分享鍵（方框加箭頭 ⬆︎）。</li><li>往下滑，按「加入主畫面」→「新增」。</li>' : '<li>按右上角的「⋮」。</li><li>按「加到主畫面」或「安裝應用程式」。</li>'}<li>之後點桌面上的「聖佳庫存」就能打開。</li></ol>`
           }
           <button class="btn block" id="j-ok" style="margin-top:12px">知道了</button>`,
          (el2, close2) => (el2.querySelector('#j-ok').onclick = close2),
        )
      }
    },
  )
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
    `<h2 class="sheet-title">共用「聖佳智慧庫存」</h2>
     <div class="form">
       <input class="field" id="sh-name" placeholder="新增使用者：名字（例如 阿明、辦公室電腦）" autocomplete="off">
       <div class="row-actions"><select class="field role-pick" id="sh-role" aria-label="權限"><option value="editor">編輯者（可以盤點、修改）</option><option value="viewer">檢視者（只能看）</option><option value="manager">管理員（也可以加人）</option></select><button class="btn" id="sh-invite">邀請</button></div>
     </div>
     <p class="section-title">擁有存取權的使用者</p>
     <div class="group" id="sh-list"><div class="row muted">載入中…</div></div>
     <p class="section-title">一般存取權</p>
     <div class="group"><div class="row"><span class="avatar lock" aria-hidden="true">${icon('lock', 18)}</span><span class="grow"><span class="title">限制</span><br><span class="meta">只有上面名單裡的人，用自己的連結碼才能開啟</span></span></div></div>
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
       .map((p) => `<button class="row" data-pick="${esc(p.id)}" ${p.id === currentId ? 'aria-current="true"' : ''}><span class="avatar" aria-hidden="true">${esc(p.name.slice(0, 1))}</span><span class="grow"><span class="title">${esc(p.name)}</span></span>${p.id === currentId ? `<span class="now-mark">${icon('check', 20)}</span>` : ''}</button>`)
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
  // 一鍵加入：同事在 LINE 點連結 → 自動打開 App → 按「加入」（不用自己複製、貼到設定）
  const link = joinLinkOf(code, appUrl)
  const text = `${name}你好：這是公司「聖佳智慧庫存」的邀請，只給你一個人用，請不要轉傳。\n\n點這個連結就會自動加入：\n${link}\n\n打開後按「加入」就好。`
  sheet(
    `<h2 class="sheet-title">邀請「${esc(name)}」</h2>
     <p class="sheet-sub">${ROLE_LABEL[role]}：${ROLE_DESC[role]}</p>
     <textarea class="field code" readonly rows="4">${esc(text)}</textarea>
     <div class="row-actions" style="margin-top:10px">${navigator.share ? '<button class="btn" id="lk-share" style="flex:1">傳給他（LINE）</button>' : ''}<button class="btn ${navigator.share ? 'secondary' : ''}" id="lk-copy" style="flex:1">複製</button></div>
     <ol class="steps-list">
       <li>私訊給「${esc(name)}」（不要貼在群組）。</li>
       <li>他在 LINE <b>點那個連結</b>，App 會自己打開，按「加入」就好。不用複製、不用到設定裡貼。</li>
       <li>這組只給他一個人用。他離職時，在「共用設定」移除權限就好，其他人不用改。</li>
     </ol>
     <p class="footnote">關掉之後就看不到這個邀請了（雲端只存雜湊值，比較安全）；忘了可以「重新產生連結碼」。</p>`,
    (el) => {
      el.querySelector('#lk-copy').onclick = async () => {
        try {
          await navigator.clipboard.writeText(text)
          toast('已複製：私訊給他')
        } catch {
          el.querySelector('textarea').select()
          toast('請長按上面的文字複製')
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
        // 共用的 AI 金鑰不會跟著搬（金鑰不從雲端傳回手機）：搬完要提醒擁有者在新位置再放一次
        const aiWas = ls.get(LS.aiShared) === '1'
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
          status(`搬家完成：大家的 App 下次同步會自動跟過來。舊的試算表和資料夾確認沒問題後可以刪掉。${aiWas ? '共用的 AI 金鑰沒有跟著搬：請到「設定 → Gemini API Key」再按一次「放到雲端」，同事才能繼續拍照辨識。' : ''}`)
          if (aiWas) {
            ls.set(LS.aiShared, '')
            alert('搬家完成。\n\n共用的 AI 金鑰沒有跟著搬（金鑰不會從雲端傳回手機）。\n請到「設定 → Gemini API Key」再按一次「放到雲端」，同事才能繼續拍照辨識。')
            render()
          } else toast('搬家完成')
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
  if (state.syncState === 'syncing') return state.syncMsg || '同步中…'
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

/** 這次盤點有沒有結果（有辨識成功的照片或手動加的）：全部失敗的不算「盤過這一格」 */
const hasResult = (s) => s.photos.some((p) => p.status === 'done') || !!(s.manual || []).length
/**
 * 同一格重盤，只算一次（跟品項庫一樣：重盤同一格以新的為準）：
 * - 有「已完成」的：用最新一次已完成的；同一格比較舊的、還沒按完成的都不算
 * - 都還沒按完成：用最新的那一次（還沒按完成的照舊算進來，畫面上會標示）
 * - 沒填位置的：沒辦法知道是不是同一格，每一次各算一筆
 * 只在你選的時間範圍裡比（今天／最近 7 天／全部）。回傳算進總表的那幾次。
 */
function countedSessions(sessions) {
  const best = new Map()
  for (const s of sessions) {
    if (!s.place || !hasResult(s)) continue
    const k = canon(s.place)
    const cur = best.get(k)
    const better = !cur || (!!s.linkedAt !== !!cur.linkedAt ? !!s.linkedAt : s.createdAt > cur.createdAt)
    if (better) best.set(k, s)
  }
  return new Set(sessions.filter((s) => (!s.place ? hasResult(s) : best.get(canon(s.place)) === s)))
}
/**
 * 把好幾次盤點合起來：同一種商品（品名＋品牌＋型號＋尺寸都一樣）的數量加總，記下在哪些位置各幾件。
 * 同一格重盤只算最新的一次（見 countedSessions）。
 * detail＝每次盤點、每一種一列（給 Excel 明細、Google 試算表用；原始紀錄，每一次都列）。
 */
function reportOf(sessions) {
  const cmp = new Intl.Collator('zh-Hant', { numeric: true }).compare
  const items = new Map()
  const detail = []
  const counted = countedSessions(sessions)
  for (const s of [...sessions].sort((a, b) => a.createdAt - b.createdAt)) {
    const place = s.place || `未填位置（${fmtTime(s.createdAt)}）`
    const use = counted.has(s)
    for (const g of groupsOf(s)) {
      const count = Number(g.count) || 0
      if (!count) continue
      detail.push({ date: ymd(s.createdAt), time: hm(s.createdAt), place: s.place || '', label: g.label, brand: g.brand, model: g.model, spec: g.spec, count, boxes: g.manual ? '' : g.boxes, source: g.manual ? '手動' : g.edited ? 'AI（有修正）' : 'AI', id: s.id, by: byName(s) })
      if (!use) continue
      const k = FIELDS.map((f) => norm(g[f])).join('|')
      if (!items.has(k)) items.set(k, { label: g.label, brand: g.brand, model: g.model, spec: g.spec, total: 0, places: new Map() })
      const it = items.get(k)
      it.total += count
      it.places.set(place, (it.places.get(place) || 0) + count)
    }
  }
  const list = [...items.values()].sort((a, b) => cmp(a.label, b.label) || cmp(a.spec, b.spec) || cmp(a.brand, b.brand))
  const places = new Set(sessions.map((s) => (s.place ? canon(s.place) : s.id)))
  const used = sessions.filter((s) => counted.has(s))
  return {
    list,
    detail,
    counted,
    total: list.reduce((n, it) => n + it.total, 0),
    places: places.size,
    sessions: sessions.length,
    // 算進總表、但還沒按「完成」的（老闆要看得出來）；同一格重盤、被比較新的取代掉的
    unfinished: used.filter((s) => !s.linkedAt).length,
    replaced: sessions.filter((s) => s.place && hasResult(s) && !counted.has(s)).length,
  }
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
      rows: [['盤點時間', '盤點人', '位置', '照片張數', '種類', '件數', '辨識模型', '按完成了沒', '算進總表'], ...[...sessions].sort((a, b) => a.createdAt - b.createdAt).map((s) => [`${ymd(s.createdAt)} ${hm(s.createdAt)}`, byName(s), s.place || '', s.photos.length, groupsOf(s).length, totalQty(s), s.model || '', s.linkedAt ? '已完成' : '還沒按完成', r.counted.has(s) ? '算' : hasResult(s) ? '不算（同一格有比較新的）' : '不算（沒有辨識成功）'])],
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
          ${
            r.unfinished || r.replaced
              ? `<div class="hint-card report-note" role="note"><div><b>總數怎麼算</b> <span class="badge new">新</span></div>${r.unfinished ? `<div><span class="badge low">含 ${r.unfinished} 次還沒按完成</span> 這幾次的數量已經算進上面，但還沒記進品項庫，可能還會改。</div>` : ''}${r.replaced ? `<div>同一格盤了好幾次的，只算一次（按過完成的優先，再看最新的）：${r.replaced} 次沒算進來。</div>` : ''}</div>`
              : ''
          }
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
          <div class="group" style="margin-top:12px"><button class="row" data-go="quality"><span class="grow"><span class="title">AI 準不準</span><br><span class="meta">AI 原本數的跟你確認後的比：數量一個不差的比例、錯在哪</span></span>${chev}</button></div>
          <p class="section-title">品項（${r.list.length}）</p>
          <div class="group${r.list.length >= 2 ? ' cols-2' : ''}">${r.list
            .map(
              (it) => `<div class="row report-row"><span class="grow"><span class="title">${esc(it.label)}</span><br><span class="meta">${esc([it.spec, it.brand, it.model].filter(Boolean).join('・') || '沒有寫尺寸')}</span>${placesHtml([...it.places])}</span><span class="report-count">${it.total}</span></div>`,
            )
            .join('')}</div>`
        : `<div class="empty"><p>${range === 'today' ? '今天還沒有盤點。' : '這段時間沒有盤點紀錄。'}<br>換一個時間範圍看看。</p></div>`
    }
  </main>`
}

// ───────────────────────── AI 準不準 ─────────────────────────
/**
 * 每次按「完成」自動算（AI 原本的數字 vs 你確認後的）；不用另外做。
 * 標準答案：兩個人各自數過、數字一樣的盤點，勾「當成標準答案」；之後可以拿那些照片重新考 AI（換模型、改設定時看有沒有變差）。
 */
async function viewQuality() {
  const all = await db.all()
  const range = state.qRange || 'all'
  const sessions = all.filter((s) => s.eval && inRange(s, range))
  const q = qualityStats(sessions)
  const items = await itemsAll()
  const golden = all.filter((s) => s.golden)
  const goldenPhotos = golden.reduce((n, s) => n + s.photos.filter((p) => p.status === 'done').length, 0)
  const runs = readJson(LS.goldenRuns, [])
  const tone = (x, good, ok) => (x == null ? '' : x >= good ? 'good' : x >= ok ? 'ok' : 'bad')
  // 最常錯的品項：同一個品項（或同名）合起來看
  const byItem = new Map()
  for (const t of sessions.flatMap((s) => s.eval.types)) {
    const k = t.item || `${t.label}|${t.spec || ''}`
    const cur = byItem.get(k) || { label: items.find((x) => x.id === t.item) ? itemTitle(items.find((x) => x.id === t.item)) : `${t.label}${t.spec ? `・${t.spec}` : ''}`, n: 0, exact: 0, err: 0 }
    cur.n += 1
    cur.exact += t.ai === t.final ? 1 : 0
    cur.err += Math.abs(t.ai - t.final)
    byItem.set(k, cur)
  }
  const worst = [...byItem.values()].filter((x) => x.exact < x.n).sort((a, b) => b.n - b.exact - (a.n - a.exact) || b.err - a.err).slice(0, 6)
  const byModel = new Map()
  for (const s of sessions) byModel.set(s.eval.model || '沒有記錄', [...(byModel.get(s.eval.model || '沒有記錄') || []), s])
  const run = state.golden
  return `
  <main class="app">
    <div class="nav">${backBtn(state.qBack || 'report', { settings: '設定', review: '盤點結果' }[state.qBack] || '總表')}</div>
    <h1 class="large-title">AI 準不準</h1>
    <p class="subtitle">拿 AI 原本數的，跟你確認後的數字比。每次按「完成」自動算，不用另外做。</p>
    <div class="seg" role="tablist" aria-label="時間範圍">${RANGES.map((x) => `<button role="tab" aria-selected="${x.id === range}" data-qrange="${x.id}">${x.label}</button>`).join('')}</div>
    ${
      q.n
        ? `<div class="stats q">
            <div class="stat ${tone(q.rate, 0.95, 0.85)}"><span class="stat-num">${pct(q.rate)}</span><span class="stat-label">數量一個不差</span></div>
            <div class="stat"><span class="stat-num">${q.mae.toFixed(1)}</span><span class="stat-label">平均差幾個</span></div>
            <div class="stat ${tone(q.wrong ? q.caught / q.wrong : null, 0.9, 0.7)}"><span class="stat-num">${q.wrong ? pct(q.caught / q.wrong) : '—'}</span><span class="stat-label">錯的有標出來</span></div>
            <div class="stat"><span class="stat-num">${q.sessions}</span><span class="stat-label">次盤點・${q.n} 種</span></div>
          </div>
          <p class="footnote" style="margin-top:8px">目標：數量一個不差 95% 以上（盤點實務的常見標準）；「錯的有標出來」越高，代表你只看「要確認」的就夠。${q.unchecked ? `<br>${q.unchecked} 種還有「要確認」沒看過就按了完成：不知道 AI 對不對，沒有算進來。` : ''}</p>
          <p class="section-title">AI 錯在哪</p>
          <div class="group">
            <div class="row"><span class="grow"><span class="title">漏數</span><br><span class="meta">你補了框，或把數量加上去（例如疊在後面看不到）</span></span><span class="qty"><b>${q.missed}</b></span></div>
            <div class="row"><span class="grow"><span class="title">多數</span><br><span class="meta">你刪掉的框，或把數量減下來（重複框、不是商品）</span></span><span class="qty"><b>${q.extra}</b></span></div>
            <div class="row"><span class="grow"><span class="title">分錯種類</span><br><span class="meta">框改成別的一種（例如 2分 看成 3分）</span></span><span class="qty"><b>${q.misclass}</b></span></div>
            <div class="row"><span class="grow"><span class="title">AI 一共框了</span><br><span class="meta">標成「要確認」的 ${q.flagged} 個，其中真的有錯 ${q.caught} 個</span></span><span class="qty"><b>${q.boxes}</b></span></div>
          </div>
          ${
            worst.length
              ? `<p class="section-title">最常數錯的</p><div class="group">${worst.map((w) => `<div class="row"><span class="grow"><span class="title">${esc(w.label)}</span><br><span class="meta">${w.n} 次裡 ${w.n - w.exact} 次數量不對，一共差 ${w.err} 個</span></span></div>`).join('')}</div>
                 <p class="footnote">常錯的那一種，拍一張清楚的樣品照（設定 → 樣品照），名稱和尺寸寫清楚，AI 下次就會拿來比。</p>`
              : ''
          }
          ${
            byModel.size > 1
              ? `<p class="section-title">各模型</p><div class="group">${[...byModel]
                  .map(([m, ss]) => {
                    const x = qualityStats(ss)
                    return `<div class="row"><span class="grow"><span class="title">${esc(m)}</span><br><span class="meta">${x.sessions} 次盤點・${x.n} 種・平均差 ${x.mae.toFixed(1)} 個</span></span><span class="qty"><b>${pct(x.rate)}</b></span></div>`
                  })
                  .join('')}</div>`
              : ''
          }`
        : `<div class="empty"><p>${range === 'all' ? '還沒有資料。' : '這段時間沒有資料。'}<br>從這一版開始，盤點完按「完成・記進品項庫」，就會開始累積。</p>${canEdit() && hasAi() && range === 'all' ? '<button class="btn small" data-action="new">開始盤點</button>' : ''}</div>`
    }
    <p class="section-title">標準答案考 AI</p>
    <div class="group">
      <div class="row"><span class="grow"><span class="title">標準答案：${golden.length} 次盤點、${goldenPhotos} 張照片</span><br><span class="meta">兩個人各自數過、數字一樣的盤點，在那次盤點最下面勾「當成標準答案」。建議湊到 30～50 張，要有難的：光線暗、疊在一起、2分和 3分 混在一起。</span></span></div>
      ${
        run
          ? `<div class="row" role="status"><span class="grow"><span class="title">AI 考試中…</span><br><span class="meta">${run.done} / ${run.total} 張照片</span></span></div>`
          : `<button class="row edit-only" data-action="golden-run" ${goldenPhotos ? '' : 'disabled'}><span class="grow"><span class="title" style="color:var(--tint)">用標準答案考 AI</span><br><span class="meta">${goldenPhotos ? `重新辨識 ${goldenPhotos} 張照片，跟標準答案比（會用到 API 額度）` : '先勾幾次「當成標準答案」才能考'}</span></span>${chev}</button>`
      }
      ${runs
        .slice(0, 6)
        .map((r) => `<div class="row"><span class="grow"><span class="title">${fmtTime(r.at)}・${esc(r.model || '')}</span><br><span class="meta">${r.photos} 張照片、${r.types} 種・平均差 ${Number(r.mae).toFixed(1)} 個${r.failed ? `・${r.failed} 張沒辨識成功` : ''}${r.worst?.length ? `<br>差最多：${esc(r.worst.join('、'))}` : ''}</span></span><span class="qty"><b>${pct(r.rate)}</b></span></div>`)
        .join('')}
    </div>
    <p class="footnote">考試時，從標準答案那幾次盤點存的樣品照不會送給 AI（不然等於先看答案）。換了模型或改了設定，考一次就知道有沒有變差。</p>
    <details class="steps"><summary>這些數字怎麼算？</summary>
      <ol>
        <li><b>數量一個不差：</b>每一種商品，AI 原本數的跟你確認後的一樣，就算一個不差。整組改名不算錯，數量對就算對。</li>
        <li><b>平均差幾個：</b>每一種差幾個（多或少都算），平均起來。</li>
        <li><b>錯的有標出來：</b>AI 錯的框（被你刪掉、改種類的）裡，有幾成事先被標成「要確認」。</li>
        <li>手動新增的種類（沒有框）不算；這一版以前的盤點沒有記 AI 原本的答案，所以不算。</li>
      </ol>
    </details>
  </main>`
}

/** 用標準答案考 AI：那些照片重新辨識一次（不看你改過的），每一種比數量 */
async function runGolden() {
  if (state.golden) return
  if (!hasAi()) return toast('要先在設定貼上 Gemini API Key，或請擁有者把金鑰放到雲端共用')
  const sessions = (await db.all()).filter((s) => s.golden)
  const photos = sessions.flatMap((s) => s.photos.filter((p) => p.status === 'done').map((p) => ({ s, p })))
  if (!photos.length) return toast('先勾幾次「當成標準答案」才能考')
  if (!confirm(`要用 AI 重新辨識 ${photos.length} 張照片（大約 ${Math.max(1, Math.round((photos.length * 15) / 60))} 分鐘），跟標準答案比。會用到 API 額度。開始嗎？`)) return
  const ids = new Set(sessions.map((s) => s.id))
  // 從標準答案那幾次存的樣品照不送（不然等於先看答案）
  const samples = (await idb.samples.all().catch(() => [])).filter((x) => !ids.has(x.sid)).slice(0, MAX_SAMPLES)
  const refs = await Promise.all(samples.map(async (x) => ({ label: x.label, brand: x.brand, model: x.model, spec: x.spec, data: await blobToBase64(x.blob) })))
  const items = await itemsAll()
  const idOf = (f) => findItem(items, f)?.id || `k:${looseKey(f)}`
  const nameOf = new Map()
  state.golden = { done: 0, total: photos.length }
  keepAwake()
  render()
  const types = []
  let model = ''
  let failed = 0
  try {
    for (const s of sessions) {
      const truth = new Map()
      for (const g of groupsOf(s).filter((x) => !x.manual)) {
        const id = s.itemOf?.[g.key] || idOf(g)
        truth.set(id, (truth.get(id) || 0) + (Number(g.count) || 0))
        if (!nameOf.has(id)) nameOf.set(id, `${g.label}${g.spec ? `・${g.spec}` : ''}`)
      }
      const guess = new Map()
      for (const p of s.photos.filter((x) => x.status === 'done')) {
        try {
          const r = await analyze({ ...p, objects: [] }, () => {}, refs)
          model = r.model || model
          for (const o of r.objects) {
            const id = idOf(o)
            guess.set(id, (guess.get(id) || 0) + 1)
            if (!nameOf.has(id)) nameOf.set(id, `${o.label}${o.spec ? `・${o.spec}` : ''}`)
          }
        } catch {
          failed += 1
        }
        state.golden.done += 1
        if (state.view === 'quality') render()
      }
      for (const id of new Set([...truth.keys(), ...guess.keys()])) types.push({ label: nameOf.get(id) || '', ai: guess.get(id) || 0, final: truth.get(id) || 0 })
    }
  } finally {
    state.golden = null
    releaseWakeLock()
  }
  const exact = types.filter((t) => t.ai === t.final).length
  const result = {
    at: Date.now(),
    model,
    photos: photos.length,
    types: types.length,
    rate: types.length ? exact / types.length : null,
    mae: types.length ? types.reduce((a, t) => a + Math.abs(t.ai - t.final), 0) / types.length : 0,
    failed,
    worst: types
      .filter((t) => t.ai !== t.final)
      .sort((a, b) => Math.abs(b.ai - b.final) - Math.abs(a.ai - a.final))
      .slice(0, 3)
      .map((t) => `${t.label}（AI ${t.ai}、答案 ${t.final}）`),
  }
  ls.set(LS.goldenRuns, JSON.stringify([result, ...readJson(LS.goldenRuns, [])].slice(0, 12)))
  if (state.view === 'quality') render()
  toast(`考完了：數量一個不差 ${pct(result.rate)}（${types.length} 種）`)
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
  // 「已刪掉第 N 張照片・復原」的提示拿掉：開始辨識後不能再把照片插回去（順序會對錯）
  document.querySelector('.toast')?.remove()
  // 佇列放照片本身（不是第幾張）：中途照片清單有變動，也不會辨識到別張
  const photos = indices.map((i) => s.photos[i]).filter(Boolean)
  state.cancel = false
  state.progress = { done: 0, total: photos.length, started: Date.now() }
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
  const queue = [...photos]
  const worker = async () => {
    while (queue.length && !state.cancel) {
      const photo = queue.shift()
      if (!s.photos.includes(photo)) continue
      // 辨識過程（用哪個模型、花幾秒、哪一步失敗）記在照片上，結果頁可以展開看、截圖回報
      photo.trace = []
      const t0 = Date.now()
      const log = (msg) => {
        photo.trace.push(`${Math.round((Date.now() - t0) / 100) / 10} 秒｜${msg}`)
        status(msg)
      }
      try {
        const r = await analyze(photo, log, refs)
        Object.assign(photo, { objects: r.objects, note: r.note, status: 'done', error: '', errorDetail: '', aiGone: [] })
        markAi(photo) // 記下 AI 原本的答案（算準確率用）
        s.model = r.model
        // 拍到儲位標籤：沒填位置就自動填（有建立過的儲位用那個寫法）
        if (r.location && !s.place) {
          s.place = findLocation(r.location)?.code || r.location.slice(0, 20).toUpperCase()
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
  if (!state.cancel) backgroundRefine(s, photos, refs)
}

/**
 * 第二輪在背景跑：同一類、但大小不一樣的框，切小圖並排再比一次。
 * 先記住要比的那幾個框（物件本身），等待時你刪了或改了框也不會對錯位置；你改過的框不會被蓋掉。
 */
async function backgroundRefine(s, photos, refs) {
  const jobs = photos
    .filter((p) => p?.status === 'done' && s.photos.includes(p))
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
      // 比對也是 AI 的答案：還沒被你改過的框，更新成比對後的種類
      markAi(photo, free.filter((o) => !o.edited && photo.objects.includes(o)))
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
  const t = e.target.closest('[data-go],[data-action],[data-open],[data-open-doubts],[data-remove-photo],[data-preview-photo],[data-photo-index],[data-focus],[data-step],[data-edit],[data-retry],[data-range],[data-qrange],[data-recount-list],[data-pick],[data-item-open],[data-item-filter],[data-place],[data-loc-edit],[data-stock-del],[data-safety],[data-example],[data-erp-cat],[data-erp-no]')
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
  if (d.go === 'quality' && state.view !== 'quality') state.qBack = ['settings', 'review'].includes(state.view) ? state.view : 'report'
  if (d.go) {
    // 新盤點拍了照片、還沒按「開始辨識」就要離開：照片只在記憶體裡，離開就不見 → 先問
    if (photosAtRisk() && d.go !== 'capture') {
      const n = state.session.photos.length
      if (!confirm(`還沒按「開始辨識」：要放棄這 ${n} 張照片嗎？\n\n放棄的話照片不會留下來。要留著就按「取消」，再按下面的「開始辨識」。`)) return
      for (const p of state.session.photos) dropUrl(p.id)
      state.session = null
    }
    // 點貨單那一頁要離開：「改一下」改了還沒按完成修改 → 先問要不要放棄；空的草稿順手刪掉
    if (state.view === 'pick-edit' && d.go !== 'pick-edit' && !(await leavePickEdit())) return
    // 首頁「該叫貨」方塊：到品項，直接開「叫貨」那一頁
    if (d.itemFilter) state.itemFilter = d.itemFilter
    return go(d.go)
  }
  // 認識產品：點類別、點產品
  if (d.erpCat !== undefined) {
    state.erp = { ...(state.erp || {}), cat: d.erpCat || null, q: '', more: 0 }
    return render()
  }
  if (d.erpNo) return erpProductSheet(d.erpNo)
  if (d.range) {
    state.range = d.range
    return render()
  }
  if (d.qrange) {
    state.qRange = d.qrange
    return render()
  }
  if ('recountList' in d) {
    state.itemFilter = 'recount'
    return go('items')
  }
  // 多選：點一筆＝勾／取消
  if (d.pick) {
    if (state.selected.has(d.pick)) state.selected.delete(d.pick)
    else state.selected.add(d.pick)
    return render()
  }
  if (d.open || d.openDoubts) {
    const s = await db.get(d.open || d.openDoubts)
    if (!s) return toast('找不到這次盤點（可能在別台被刪掉了）')
    state.session = s
    delete state.session.edits // 第一版的舊欄位，不再使用
    await itemsAll()
    go('review', { photoIndex: 0 })
    // 照片還沒下載的：現在抓這一次的（畫面先出來，照片到了再補上）
    fetchPhotos(state.session).catch((e) => toast(`照片下載不了：${e.message}`))
    // 首頁「要確認」方塊：直接跳出「這一個一樣嗎？」，不用再按一次
    if (d.openDoubts && canEdit()) {
      const list = doubtsOf(state.session)
      if (list.length) quickSheet(list, 0, { doubt: true })
    }
    return
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
    // 刪照片：先拿掉，提示可以「復原」（以前點縮圖就直接刪、沒有確認）
    const s = state.session
    const i = Number(d.removePhoto)
    const [ph] = s.photos.splice(i, 1)
    if (!ph) return
    // 連續刪好幾張：按一次「復原」全部放回原來的位置
    if (removedPhotos.s !== s) removedPhotos = { s, list: [], timer: 0 }
    removedPhotos.list.push({ ph, i })
    render()
    // 復原時間過了還是沒放回去：放掉這些照片占的記憶體
    clearTimeout(removedPhotos.timer)
    const batch = removedPhotos
    batch.timer = setTimeout(() => {
      for (const r of batch.list) if (!s.photos.includes(r.ph)) dropUrl(r.ph.id)
      if (removedPhotos === batch) removedPhotos = { s: null, list: [], timer: 0 }
    }, 7000)
    const n = batch.list.length
    return toast(n > 1 ? `已刪掉 ${n} 張照片` : `已刪掉第 ${i + 1} 張照片`, {
      label: '復原',
      run: () => {
        // 已經開始辨識（或換了一次盤點）就不能放回去：辨識是照順序排好的，插回去會對錯照片
        if (state.session !== s || state.view !== 'capture' || s.photos.some((p) => p.status !== 'pending')) return toast('已經開始辨識了，沒辦法復原；要的話請重拍')
        // 倒著放回去：每一張都回到刪掉前的位置
        for (const r of [...batch.list].reverse()) s.photos.splice(Math.min(r.i, s.photos.length), 0, r.ph)
        batch.list = []
        render()
        toast(n > 1 ? `已放回 ${n} 張照片` : '已放回這張照片')
      },
    })
  }
  if (d.previewPhoto !== undefined) {
    const s = state.session
    const i = Number(d.previewPhoto)
    const p = s?.photos[i]
    if (!p) return
    // 跟全站一樣用 photoViewer：可以兩指放大、左右換張；下面多一顆「刪掉這張」
    const list = s.photos.map((x, k) => ({ src: urlOf(x), cap: `第 ${k + 1} 張照片・按「開始辨識」才會送給 AI`, k }))
    return photoViewer(list, i, {
      action: {
        label: '刪掉這張',
        icon: 'trash',
        cls: 'danger',
        run: async (en, close) => {
          close()
          document.querySelector(`[data-remove-photo="${en.k}"]`)?.click()
          // 原本的縮圖刪掉了：焦點放到同一個位置的下一張（沒有了就前一張）
          await new Promise((r) => setTimeout(r, 60))
          const left = state.session?.photos.length || 0
          if (left) $app.querySelector(`[data-preview-photo="${Math.min(en.k, left - 1)}"]`)?.focus({ preventScroll: true })
        },
      },
    })
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
    // ── 首頁多選 ──
    case 'select-start':
      state.selecting = true
      state.selected = new Set()
      return render()
    case 'select-done':
      state.selecting = false
      return render()
    case 'select-all': {
      const all = await db.all()
      state.selected = state.selected.size === all.length ? new Set() : new Set(all.map((s) => s.id))
      return render()
    }
    case 'select-unfinished':
      state.selected = new Set((await db.all()).filter((s) => !s.linkedAt).map((s) => s.id))
      return render()
    case 'bulk-export': {
      const list = (await db.all()).filter((s) => state.selected.has(s.id))
      if (!list.length) return
      downloadBlob(reportXlsx(list, await itemsAll(true), await whoOfStock()), `盤點_${list.length}次_${ymd(Date.now())}.xlsx`)
      return toast(`已下載 Excel（${list.length} 次盤點）`)
    }
    case 'bulk-finish': {
      const list = (await db.all()).filter((s) => state.selected.has(s.id) && s.photos.some((p) => p.status === 'done'))
      if (!list.length) return toast('選到的都沒有辨識成功的照片，沒辦法記進品項庫')
      // 沒填位置的，每一次各算一筆：同一批貨盤兩次就會重複算
      const noPlace = list.filter((s) => !s.place).length
      if (!confirm(`把 ${list.length} 次盤點記進品項庫？${noPlace ? `\n\n其中 ${noPlace} 次沒填位置：沒填位置的每一次都會各算一筆，同一批貨盤過兩次會重複算。測試用的請先刪掉。` : ''}`)) return
      t.disabled = true
      let created = 0
      let failed = 0
      for (const s of list) {
        try {
          created += (await linkSession(s)).created
        } catch {
          failed += 1
        }
      }
      state.selecting = false
      toast(`已記進品項庫：${list.length - failed} 次盤點${created ? `（新的品項 ${created} 種，到「品項」確認名稱）` : ''}${failed ? `；${failed} 次沒成功，請打開那次再按完成` : ''}`)
      render()
      if (syncReady()) syncNow().catch((err) => toast(`同步沒成功：${err.message}；有網路時會再自動試`))
      if (ls.get(LS.sheet) && ls.get(LS.autoSync, '1') === '1') syncToSheet(list).catch((err) => toast(`試算表同步沒成功：${err.message}`))
      return
    }
    case 'bulk-delete': {
      const list = (await db.all()).filter((s) => state.selected.has(s.id))
      // 編輯者一次最多刪 20 次（雲端也是這樣擋：刪更多＝刪光大家的紀錄，要擁有者或管理員來）
      if (!canManage() && syncReady() && list.length > 20) return toast(`一次最多刪 20 次盤點（你選了 ${list.length} 次）；要刪更多請找擁有者或管理員`)
      if (!list.length || !confirm(`刪除 ${list.length} 次盤點？照片和結果都會刪掉，記在品項庫的數量也會拿掉。刪掉的找不回來。`)) return
      t.disabled = true
      for (const s of list) {
        await db.del(s.id)
        tombstone(`session:${s.id}`)
        await unlinkSession(s.id)
      }
      state.selecting = false
      toast(`已刪除 ${list.length} 次盤點`)
      render()
      if (syncReady()) syncNow().catch(() => {})
      return
    }
    case 'save-key': {
      const v = document.getElementById('apikey').value.trim()
      ls.set(LS.key, v)
      ls.set(LS.model, '')
      if (!v) return toast('已清除 API Key')
      const done = busyBtn(t, '測試中…')
      try {
        const list = await fetchModels()
        if (!list.length) return toast('Key 可以用，但找不到能看圖的模型')
        ls.set(LS.model, list[0])
        toast(`可以用了！模型：${list[0]}`)
        return render()
      } catch (err) {
        return toast(err.message)
      } finally {
        done()
      }
    }
    case 'toggle-key': {
      const el = document.getElementById('apikey')
      el.type = el.type === 'password' ? 'text' : 'password'
      return
    }
    case 'ai-share': {
      // 擁有者：把金鑰放到自己的 Apps Script（指令碼屬性），大家共用；同事手機上不會有金鑰
      // 只放「存好、測過」的金鑰：輸入框裡還沒按「儲存並測試」的不算（打錯一個字，大家都不能用）
      const v = ls.get(LS.key).trim()
      const typed = (document.getElementById('apikey')?.value || '').trim()
      if (!v || (typed && typed !== v)) return toast('先在上面貼上 API Key、按「儲存並測試」確認可以用，再放到雲端')
      if (!confirm('把這把金鑰放到你的 Apps Script（雲端）給大家共用？\n同事的手機不會拿到金鑰；大家用的次數都算在這把金鑰上（每人每天最多 300 次）。')) return
      const done = busyBtn(t, '測試中…')
      try {
        // 放上去之前再測一次（金鑰可能已經被停用）
        try {
          await fetchModels()
        } catch (err) {
          return toast(`這把金鑰現在不能用，沒有放到雲端：${err.message}`)
        }
        const r = await postSync({ action: 'aiKey', aiKey: v })
        if (!('ai' in r)) return toast('雲端程式碼是舊版：請先到「Google 試算表」複製試算表程式碼、貼上、部署新版本')
        ls.set(LS.aiShared, '1')
        toast('已放到雲端：同事打開 App 就能拍照辨識，不用再跟你要金鑰')
        return render()
      } catch (err) {
        return toast(err.message)
      } finally {
        done()
      }
    }
    case 'ai-unshare': {
      if (!confirm('停止共用金鑰？之後同事要拍照辨識，就得自己貼金鑰。')) return
      try {
        await postSync({ action: 'aiKey', aiKey: '' }) // 雲端說不行（例如不是擁有者）會丟錯誤，下面會顯示
        ls.set(LS.aiShared, '')
        toast('已停止共用金鑰')
        return render()
      } catch (err) {
        return toast(err.message)
      }
    }
    case 'lookup-no-ai':
      return toast(!canEdit() && aiSharedOn() ? '檢視者不能用公司共用的金鑰拍標籤；打型號查詢可以用' : syncReady() ? '先請擁有者到「設定」把 AI 金鑰放到雲端，或到「設定」貼上自己的 Gemini API Key' : '先到「設定」貼上 Gemini API Key，才能拍標籤；打型號查詢不用')
    case 'pick-model':
      if (!hasAi()) return toast(canEdit() ? '先在上面貼上 Gemini API Key' : '檢視者不能用公司共用的金鑰')
      try {
        const list = await fetchModels()
        const rows = (arr) => arr.map((m) => `<button class="row" data-action="use-model" data-model="${esc(m)}"><span class="grow">${esc(m)}</span>${m === ls.get(LS.model) ? `<span class="now-mark">${icon('check', 20)}</span>` : ''}</button>`).join('')
        document.getElementById('models').innerHTML = `<p class="section-title">快速（Flash）</p><div class="group">${rows(list.slice(0, 6))}</div><p class="footnote">排越前面越推薦（新、穩定、快）。平常用這個。</p>${
          proCache.length
            ? `<p class="section-title">精細（Pro）</p><div class="group">${rows(proCache.slice(0, 3))}</div><p class="footnote">看得比較細，相似品比較分得開；但每張要等比較久，免費額度也比 Flash 少很多。Flash 分不出來、又不想存樣品照時再試。</p>`
            : ''
        }`
        glueTails(document.getElementById('models'))
      } catch (err) {
        toast(err.message)
      }
      return
    case 'diagnose': {
      if (!hasAi()) return toast(canEdit() ? '先在上面貼上 Gemini API Key' : '檢視者不能用公司共用的金鑰')
      const out = document.getElementById('diag')
      const lines = []
      const show = () => (out.innerHTML = `<div class="group">${lines.map((l) => `<div class="row"><span class="grow" style="font-size:14px;word-break:break-all">${l}</span></div>`).join('')}</div>`)
      lines.push('版本 ' + esc(VERSION))
      lines.push('API Key：' + (ls.get(LS.key) ? '有（' + esc(ls.get(LS.key).slice(0, 6)) + '…）' : aiViaSheet() ? '用公司共用的（放在雲端）' : '<b>沒有</b>'))
      show()
      try {
        const list = await fetchModels()
        lines.push(`<span class="diag-ok">${icon('check', 16)}</span>模型清單：找到 ` + list.length + ' 個可用（' + esc(list.slice(0, 3).join('、')) + '）')
        show()
        const model = await currentModel()
        const t0 = Date.now()
        const data = await call('models/' + model + ':generateContent', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: '只回答 OK' }] }] }) })
        const reply = data?.candidates?.[0]?.content?.parts?.map((p) => p.text || '').join('') || '（沒有回覆）'
        lines.push(`<span class="diag-ok">${icon('check', 16)}</span>` + esc(model) + ' 回覆「' + esc(reply.trim().slice(0, 20)) + '」，花 ' + ((Date.now() - t0) / 1000).toFixed(1) + ' 秒')
      } catch (err) {
        lines.push(`<span class="diag-bad">${icon('x', 16)}</span>` + esc(err.message) + (err.detail ? '<br><span class="muted">' + esc(err.detail) + '</span>' : ''))
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
      const raw = document.getElementById('link-code').value.trim()
      const code = parseLinkCode(raw)
      // 擁有者常把自己的試算表網址貼到這格（沒有 #k=…）：直接告訴他該按哪裡
      if (!code && /^https:\/\/script\.google\.com\/macros\/s\/[^#\s]+\/exec$/.test(raw)) {
        if (!ls.get(LS.sheet)) {
          ls.set(LS.sheet, raw)
          render()
        }
        return toast('這是試算表網址（擁有者用），不是連結碼：請按上面「開啟多人同步（我是擁有者）」。連結碼是同事收到的，後面會有 #k=…')
      }
      if (!code) return toast('連結碼不對：要整段貼上（收到的那一整段網址）')
      try {
        await joinWith(code)
        toast('連結成功，下載資料中…')
      } catch (err) {
        toast(err.message)
      }
      return
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
      // 已經在同步了：不要再開一次，也不要讓人以為壞掉
      if (syncing) return toast(`已經在同步了，請等一下（${state.syncMsg || '照片多會比較久'}）`)
      toast('同步中…')
      return syncNow((m) => toast(m))
        .then((r) => toast(r.pushed || r.pulled ? `同步完成：上傳 ${r.pushed}、下載 ${r.pulled}` : '已經是最新的'))
        .catch((err) => toast(`同步沒成功：${err.message}`))
        .finally(() => ['home', 'settings'].includes(state.view) && render())
    case 'finish': {
      // 完成：記進品項庫；有設定 Google 試算表又開著自動同步，就順便寫進去（在背景送，不用等）
      const s = state.session
      // 還有「要確認」沒看過：先提醒（按取消＝回去一個一個確認；按確定＝還是完成，沒確認的不算進「AI 準不準」）
      const open = s ? doubtsOf(s).length : 0
      if (open && !confirm(`還有 ${open} 個「要確認」沒看過。\n\n按「取消」回去一個一個確認；\n按「確定」還是完成（沒確認的不會算進「AI 準不準」）。`)) return quickSheet(doubtsOf(s), 0, { doubt: true })
      if (s) {
        // 記進品項庫要 1～2 秒：按鈕先變「記進品項庫中…」，不然像沒反應
        busyBtn(t, '記進品項庫中…')
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
    case 'golden-run':
      return runGolden().catch((err) => {
        state.golden = null
        toast(`考試沒有完成：${err.message || err}`)
        render()
      })
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
      const done = busyBtn(t, '測試中…')
      try {
        const data = await (await fetch(url)).json()
        toast(data.ok ? `連上了：${data.sheet || 'Google 試算表'}` : '連得到，但回覆不對；請確認貼的是這個 App 的程式碼')
      } catch {
        toast('已儲存；測試時瀏覽器讀不到回覆，同步一次後打開試算表確認')
      } finally {
        done()
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
    case 'repair-photos':
      if (!syncReady()) return toast('這台沒有開啟多台同步，雲端沒有備份：照片拿不回來')
      toast('正在從雲端拿照片…')
      // 新存法先抓；沒有的話（舊格式、照片跟資料綁一起）再從頭下載一次
      return fetchPhotos(state.session, { quiet: true })
        .then((n) => (n ? toast(`拿回 ${n} 張照片`) : fullPull()))
        .catch((e) => toast(`照片沒拿回來：${e.message}`))
    case 'full-pull':
      return fullPull().catch((e) => toast(`沒有下載成功：${e.message}`))
    case 'add-box':
      state.addMode = true
      state.focus = null
      render()
      return toast('點照片上漏掉的那一個')
    case 'add-cancel':
      state.addMode = false
      return render()
    case 'zoom':
      return openReviewViewer()
    case 'zoom-in':
      state.zoom = clampZoom(Math.floor(state.zoom) + 1)
      return render()
    case 'zoom-out':
      state.zoom = clampZoom(Math.ceil(state.zoom) - 1)
      return render()
    case 'zoom-close':
      return closeReviewViewer()
    case 'del-sample':
      if (!confirm('刪掉這張樣品照？')) return
      await idb.samples.del(d.id)
      tombstone(`sample:${d.id}`)
      toast('已刪掉樣品照')
      return render()
    case 'clear-all':
      // 有同步時會刪掉每一台的紀錄：只給擁有者、管理員（雲端也擋）
      if (!canManage()) return toast('只有擁有者或管理員可以刪除全部盤點紀錄')
      if (!confirm(`確定刪除全部盤點紀錄（含照片）？刪了救不回來。\n品項庫、各位置的數量、儲位會保留。${syncReady() ? '\n有開多台同步：其他裝置的盤點紀錄也會一起刪掉。' : ''}`)) return
      for (const s of await db.all()) tombstone(`session:${s.id}`)
      await db.clear()
      return toast('已全部刪除；品項庫保留')
    // ── 品項庫 ──
    case 'item-add':
      return itemEditSheet(null)
    case 'items-more':
      return itemsMoreSheet()
    case 'side-toggle':
      ls.set(LS.sideCollapsed, sideCollapsed() ? '0' : '1')
      document.body.classList.toggle('side-collapsed', sideCollapsed())
      return render()
    case 'erp-import':
      if (!canManage()) return toast('匯入正航產品表會改大家的帳面數：只有擁有者或管理員可以匯入')
      return erpImportSheet()
    case 'erp-cats':
      if (!canManage()) return toast('類別名稱由擁有者或管理員設定')
      return erpCatsSheet()
    case 'erp-stocked':
      state.erp.stocked = !state.erp.stocked
      state.erp.more = 0
      return render()
    case 'erp-more':
      state.erp.more += 150
      return render()
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
      return numberSheet({ title: inbound ? '進貨幾個？' : '賣出幾個？', sub: `${esc(itemTitle(it))}：帳面數${inbound ? '加' : '減'}這麼多。${blindMe() ? '' : it.book == null ? `還沒有帳面數，會從實盤 ${onHand(it)} 開始算。` : `目前帳面 ${it.book}。`}`, value: '1', action: inbound ? '記進貨' : '記賣出' }, async (n) => {
        const now = Date.now()
        // 還沒有帳面數：先記一筆「從實盤開始」，帳面數才算得回來（多人同步時用進出紀錄重算）
        if (it.book == null) it.moves.push({ at: now - 1, kind: 'set', qty: onHand(it) })
        it.moves.push({ at: now, kind: inbound ? 'in' : 'out', qty: n })
        it.book = bookFromMoves(it.moves)
        await putItem(it)
        toast(blindMe() ? (inbound ? `已記進貨 ${n} 個` : `已記賣出 ${n} 個`) : `帳面數變成 ${it.book}`)
      })
    }
    case 'recount': {
      const it = await currentItem()
      const rc = it?.recounts?.[d.sid]
      if (!rc || rc.status !== 'pending') return render()
      const me = currentCounter() || { id: ls.get(LS.memberId) || '', name: ls.get(LS.memberName) || '' }
      const same = me.id && me.id === rc.first.byId && roster().length > 1
      // 複盤一律看不到第一次的數字（不然會照著數）
      return numberSheet({ title: `${rc.place ? placeLabel(rc.place) : '這一格'} 有幾個？`, titleHtml: `${rc.place ? placeHtml(rc.place, { name: false }) : '這一格'} 有幾個？`, sub: `${esc(itemTitle(it))}：到那一格自己數一次。${same ? '<br><b>這一格是你第一次盤的，最好請另一個人複盤。</b>' : ''}`, value: '', action: '記下複盤數字' }, async (n) => {
        const now = Date.now()
        const earlier = [rc.first.count, ...rc.counts.map((c) => c.count)]
        rc.counts.push({ count: n, by: me.name, byId: me.id, at: now })
        // 這一格改成複盤的數字（同一次盤點，時間記成現在）
        const st = it.stock?.[rc.pk]
        // rc：這一格是複盤的數字（重按那次盤點的「完成」不會蓋回第一次的數字；之後新的盤點才會取代）
        if (st && !st.removed) it.stock[rc.pk] = { ...st, count: n, at: Math.max(st.at || 0, now), v: now, rc: true }
        if (!diffOf(it)) Object.assign(rc, { status: 'fixed', reason: REASONS[0] })
        else rc.status = earlier.includes(n) ? 'confirmed' : 'pending'
        rc.v = now
        await putItem(it)
        if (rc.status === 'fixed') return toast('跟帳面一樣了：第一次數錯，已經更正')
        if (rc.status === 'pending') return toast('跟前一次數的不一樣：已改成這次的數字，請再找一個人數一次')
        toast('兩次數字一樣，差異確定：請選原因')
        reasonSheet(it, d.sid)
      })
    }
    case 'recount-reason': {
      const it = await currentItem()
      if (it?.recounts?.[d.sid]) reasonSheet(it, d.sid)
      return
    }
    case 'recount-adjust':
    case 'recount-keep': {
      if (!canManage()) return toast('只有擁有者或管理員可以決定要不要調整帳面')
      const it = await currentItem()
      const rc = it?.recounts?.[d.sid]
      if (!rc || rc.status !== 'confirmed') return render()
      const now = Date.now()
      if (d.action === 'recount-adjust') {
        if (!confirm(`帳面數改成實盤 ${onHand(it)} 個？${rc.reason ? `（原因：${rc.reason}）` : ''}`)) return
        it.moves.push({ at: now, kind: 'set', qty: onHand(it), why: rc.reason || '複盤確定' })
        it.book = bookFromMoves(it.moves)
      }
      Object.assign(rc, { status: d.action === 'recount-adjust' ? 'adjusted' : 'kept', doneBy: ls.get(LS.memberName) || '', doneAt: now, v: now })
      await putItem(it)
      toast(d.action === 'recount-adjust' ? `帳面數改成 ${it.book}` : '保留帳面數，不調整')
      return render()
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
        numberSheet({ title: '剩幾個就要叫貨？', sub: `${esc(itemTitle(it))}：${blindMe() ? `實盤 ${onHand(it)} 個` : esc(orderWhy(it))}。填 3 的意思是：剩 3 個以下就提醒叫貨（實盤、帳面看比較少的）。`, value: it.safety ?? 3, action: '儲存' }, async (n) => {
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
    // ── 點貨對單（4.7 測試版） ──
    case 'pk-new': {
      // 先不存：加了品項或填了單號才存（按進來又返回，不會留下一張空的）
      const c = currentCounter()
      state.pick = { id: uid(), kind: d.kind === 'in' ? 'in' : 'out', ref: '', createdAt: Date.now(), by: c?.name || '', byId: c?.id || '', lines: [] }
      state.pickEdit = null
      state.pickQ = ''
      return go('pick-edit')
    }
    case 'pk-open': {
      const p = await idb.picks.get(d.id)
      if (!p) return toast('找不到這筆點貨紀錄（可能在別台被刪掉了）')
      state.pick = p
      state.pickEdit = null
      state.pickQ = ''
      return go(!p.doneAt && canEdit() ? 'pick-edit' : 'pick-result')
    }
    case 'pk-more-list':
      state.pickMore = (state.pickMore || 0) + 30
      return render()
    case 'pk-add': {
      const ix = await pickCtx()
      const r = d.no ? rowOfNo(ix, d.no) : ix.itemById.get(d.iid) ? { it: ix.itemById.get(d.iid) } : null
      if (!r) return toast('找不到這個產品，請重新搜尋')
      return addPickLine(r)
    }
    case 'pk-step': {
      const l = pickLine(d.id)
      if (!l) return
      l.got = Math.max(0, Math.min(PICK_QTY_MAX, round2((gotOf(l) ?? 0) + Number(d.d || 0))))
      paintLine(l)
      return pickSave()
    }
    case 'pk-done': {
      const l = pickLine(d.id)
      if (!l) return
      const st = lineState(l)
      // 琥珀色（對不到、AI 配的還沒確認）：先選產品／確認，才能點好了
      if (st === 'none' || st === 'guess') {
        flashLine(l.id, { scroll: false })
        // ⁠（不換行的連接字）：「是這個」「換一個」不會被拆成兩行
        return toast(st === 'guess' ? '先確認是不是這個產品：按「是⁠這⁠個」或「換⁠一⁠個」，才能點好了' : '先選是哪一個產品，才能點好了')
      }
      // 已經點好了：再按不會取消（連點兩下也不會不小心取消）；要改回去用 ⋯
      if (l.done) return toast('已經點好了。要改回「還沒點」，按這一項右上的 ⋯')
      // 單子數量沒讀到（AI 看不清楚）：先補數量，再算點好了
      if (l.qty == null)
        return editLineQty(l, () => {
          if (l.got == null) l.got = l.qty
          l.done = true
        })
      // 實拿沒改＝跟單子一樣
      if (l.got == null) l.got = l.qty
      l.done = true
      paintLine(l)
      return pickSave()
    }
    case 'pk-confirm': {
      const l = pickLine(d.id)
      return l && confirmLine(l)
    }
    case 'pk-qty': {
      const l = pickLine(d.id)
      return l && editLineQty(l)
    }
    case 'pk-more': {
      const l = pickLine(d.id)
      return l && lineMoreSheet(l)
    }
    case 'pk-assign': {
      const l = pickLine(d.id)
      return l && assignLine(l)
    }
    case 'pk-remove': {
      const l = pickLine(d.id)
      return l && removeLine(l)
    }
    case 'pk-finish': {
      const p = state.pick
      if (!p) return
      if (state.pickBusyId === p.id) return toast('AI 還在讀單子，讀完再按完成')
      if (!p.lines.length) return toast(editingCopy() ? '完成的紀錄至少要留 1 項；整張都不要了，請擁有者或管理員在結果頁刪除' : '還沒有品項：先把單子上的品項加進來')
      if (p.lines.some((l) => !l.done || ['none', 'guess'].includes(lineState(l)))) return finishPickSheet(p)
      return finishPick(p)
    }
    case 'pk-discard': {
      const p = state.pick
      if (!p || editingCopy()) return go('pick')
      if (state.pickBusyId === p.id) return toast('AI 還在讀單子，讀完再刪')
      if (p.lines.length && !confirm(`刪掉這張點貨單？已經加的 ${p.lines.length} 項都會刪掉。`)) return
      p._gone = true
      if (pickSaving) await pickSaving
      await idb.picks.del(p.id).catch(() => {})
      if (p._syncT) tombstone(`pick:${p.id}`)
      picksCache = null
      state.pick = null
      toast(p.lines.length || p.ref ? '已刪掉這張點貨單' : '已刪掉（還沒有品項）')
      return go('pick')
    }
    case 'pk-cancel-edit': {
      // 放棄「改一下」：回到原本的紀錄（沒有改到任何東西）
      const id = state.pick?.id
      if (!(await leavePickEdit())) return
      const stored = id ? await idb.picks.get(id).catch(() => null) : null
      state.pick = stored
      return go(stored ? 'pick-result' : 'pick')
    }
    case 'pk-reopen': {
      // 改一下＝暫存的副本：按「完成修改」才寫回紀錄、才同步；中途離開會問要不要放棄
      const p = state.pick
      if (!p?.doneAt) return go('pick')
      state.pickEdit = { id: p.id, base: clonePick(p) }
      state.pick = { ...clonePick(p), _copy: true }
      state.pickQ = ''
      return go('pick-edit')
    }
    case 'pk-delete': {
      if (!canManage()) return toast('只有擁有者或管理員可以刪點貨紀錄')
      const p = state.pick
      if (!p || !confirm(`刪除這筆點貨紀錄（${PICK_KIND[p.kind] || '點貨'}${p.ref ? `・${p.ref}` : ''}）？${syncReady() ? '\n有開多台同步：大家的手機都會一起刪掉，' : '\n'}交叉比對也不會再算它。`)) return
      p._gone = true
      await idb.picks.del(p.id)
      tombstone(`pick:${p.id}`)
      picksCache = null
      state.pick = null
      toast('已刪除這筆點貨紀錄')
      return go('pick')
    }
    case 'pk-copy': {
      const p = state.pick
      if (!p) return
      try {
        await navigator.clipboard.writeText(pickText(p, await pickCtx()))
        return toast('已複製：到 LINE 貼上就好')
      } catch {
        return toast('這個瀏覽器不讓複製，請直接截圖')
      }
    }
    case 'pk-no-ai':
      return toast(!canEdit() && aiSharedOn() ? '檢視者不能用公司共用的金鑰' : syncReady() ? '先請擁有者到「設定」把 AI 金鑰放到雲端，或到「設定」貼上自己的 Gemini API Key；打編號、搜品名不用 AI' : '先到「設定」貼上 Gemini API Key，才能拍單子讓 AI 讀；打編號、搜品名不用 AI')
    case 'pk-cross-last': {
      if (!canManage()) return toast('交叉比對只有擁有者或管理員看得到')
      const last = readJson(LS.pickCrossLast, null)
      if (!last) return toast('還沒有比對結果：下次匯入正航產品表時會自動比對')
      return crossListSheet(last)
    }
  }
})

function bindInputs() {
  document.getElementById('cam')?.addEventListener('change', (e) => addFiles([...e.target.files]))
  document.getElementById('pick')?.addEventListener('change', (e) => addFiles([...e.target.files]))
  document.getElementById('sample-cam')?.addEventListener('change', (e) => e.target.files[0] && newSampleSheet(e.target.files[0]))
  document.getElementById('auto-sync')?.addEventListener('change', (e) => ls.set(LS.autoSync, e.target.checked ? '1' : '0'))
  document.getElementById('golden')?.addEventListener('change', async (e) => {
    const s = state.session
    s.golden = e.target.checked ? { at: Date.now(), by: currentCounter()?.name || ls.get(LS.memberName) || '' } : null
    await save()
    toast(e.target.checked ? '已當成標準答案：到「AI 準不準」可以拿來考 AI' : '已取消標準答案')
  })
  // 點貨交叉比對（測試版）：只在這台、只有擁有者／管理員能開（不同步：匯入正航的人自己決定要不要比）
  document.getElementById('rule-pick-cross')?.addEventListener('change', (e) => {
    if (!canManage()) return render()
    ls.set(LS.pickCross, e.target.checked ? '1' : '')
    toast(e.target.checked ? '已開啟：下次匯入正航產品表時會跟點貨紀錄比對（第一次只記基準）' : '已關閉：匯入正航時不比對')
    render()
  })
  bindPickInputs()
  for (const [id, k] of [
    ['rule-blind', LS.blind],
    ['rule-recount', LS.recount],
  ])
    document.getElementById(id)?.addEventListener('change', (e) => {
      if (!canManage()) return render()
      ls.set(k, e.target.checked ? '1' : '0')
      touchSettings() // 跟著共用設定同步給大家
      toast(e.target.checked ? '已開啟，會同步給大家' : '已關閉，會同步給大家')
      render()
    })
  document.getElementById('place')?.addEventListener('input', (e) => (state.session.place = e.target.value))
  document.getElementById('item-search')?.addEventListener('input', (e) => filterRows($app, e.target.value))
  // 認識產品的搜尋：只重畫清單，打字的框不會跳掉
  document.getElementById('erp-q')?.addEventListener('input', async (e) => {
    state.erp.q = e.target.value
    state.erp.more = 0
    const list = document.getElementById('erp-list')
    if (list) {
      list.innerHTML = erpListHtml(await erpGet(), await itemsAll())
      glueTails(list)
    }
  })
  document.getElementById('lookup-q')?.addEventListener('input', (e) => {
    // 只更新結果區，輸入框不重畫（打字、選字不會被打斷）
    if (state.lookup.read?.url) URL.revokeObjectURL(state.lookup.read.url)
    state.lookup = { q: e.target.value.trim(), read: null, busy: false }
    const box = document.getElementById('lookup-results')
    box.innerHTML = lookupResults()
    glueTails(box)
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
// 新版裝好了：不要馬上重新整理（拍了照片還沒辨識的話，照片只在記憶體裡，重新整理就全部不見）→ 等到安全的時候才重新整理
const photosAtRisk = () => state.view === 'capture' && !!state.session?.photos.length
const safeToReload = () =>
  // 只在清單類的頁面重新整理（盤點結果、品項頁、拍照頁等你換頁再說）
  ['home', 'items', 'lookup', 'settings', 'report', 'catalog', 'locations', 'quality', 'pick', 'pick-result'].includes(state.view) &&
  !photosAtRisk() &&
  !state.busy &&
  !state.pickBusyId &&
  !document.querySelector('.crop-wrap') &&
  !document.querySelector('.viewer.pv') && // 正在放大看照片
  !state.refining &&
  !state.lookup?.busy &&
  !state.moving &&
  !syncing &&
  !document.querySelector('.sheet-backdrop') &&
  !['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement?.tagName)
// 拍了照片還沒辨識就要關掉／重新整理分頁：瀏覽器會先問一下（iPhone 的 App 模式不一定會問）
addEventListener('beforeunload', (e) => {
  if (!photosAtRisk()) return
  e.preventDefault()
  e.returnValue = ''
})
function tryReload() {
  if (!updateReady) return
  if (safeToReload()) return location.reload()
  if (!updateTold) {
    updateTold = true
    toast('新版已經下載好：這一頁做完、換頁時會自動更新')
  }
}
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.addEventListener('message', (e) => {
    if (e.data?.type !== 'sw-updated') return
    // 回覆服務工作程式「我自己會重新整理」（舊版 App 不會回，服務工作程式就直接幫它重新整理）
    e.ports?.[0]?.postMessage({ type: 'sw-ack' })
    updateReady = true
    tryReload()
  })
  // 保險：沒收到通知、但已經換成新的服務工作程式（不是第一次安裝）→ 一樣等安全的時候重新整理
  const hadController = !!navigator.serviceWorker.controller
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (!hadController) return
    updateReady = true
    tryReload()
  })
  navigator.serviceWorker
    .register('sw.js', { updateViaCache: 'none' })
    .then((reg) => {
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState !== 'visible') return
        // 拍照拍到一半（照片還沒辨識）不要檢查更新
        if (!state.busy && state.view !== 'analyzing' && !photosAtRisk()) reg.update().catch(() => {})
        tryReload()
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
// 點邀請連結打開的（#join=…）：直接跳出「加入」
{
  const code = location.hash.startsWith('#join=') ? parseLinkCode(location.hash) : null
  if (code) setTimeout(() => joinSheet(code), 300)
}
// 打開 App：有開同步就先跟大家對一次
if (syncReady()) setTimeout(() => syncNow().catch(() => {}), 1200)
