// 讓 App 本身離線也能打開（AI 辨識還是要網路）
// 更新：每次都先問 GitHub 有沒有新版（跳過手機的 10 分鐘暫存）；新版裝好後，App 在安全的頁面（不是拍照拍到一半）自動重新整理
const CACHE = 'inventory-v45'
const SHELL = ['./', 'index.html', 'app.js', 'rules.js', 'icons.js', 'styles.css', 'manifest.webmanifest', 'icon.svg', 'apple-touch-icon.png', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'logo.svg', 'xlsx.js', 'google-sheets.gs']

self.addEventListener('install', (e) => {
  // cache: 'reload'＝直接跟 GitHub 拿最新的，不用手機暫存
  e.waitUntil(
    caches
      .open(CACHE)
      .then((c) => c.addAll(SHELL.map((u) => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting()),
  )
})

/**
 * 問一個開著的畫面：「新版好了，你自己會重新整理嗎？」
 * 4.6.1 起的 App 會回 sw-ack，自己挑安全的時候重新整理（拍照拍到一半不會把照片弄丟）；3 秒沒回＝舊版 App。
 * 用 MessageChannel 回覆：服務工作程式還在「啟用中」也收得到。
 */
const askClient = (w) =>
  new Promise((resolve) => {
    const ch = new MessageChannel()
    const timer = setTimeout(() => resolve(false), 3000)
    ch.port1.onmessage = (ev) => {
      clearTimeout(timer)
      resolve(ev.data?.type === 'sw-ack')
    }
    try {
      w.postMessage({ type: 'sw-updated', cache: CACHE }, [ch.port2])
    } catch {
      clearTimeout(timer)
      resolve(false)
    }
  })

self.addEventListener('activate', (e) => {
  e.waitUntil(
    (async () => {
      const keys = await caches.keys()
      const hadOld = keys.some((k) => k !== CACHE)
      await Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
      await self.clients.claim()
      // 從舊版升上來（第一次安裝不用）：通知開著的畫面「新版好了」，讓 App 自己找時間重新整理；
      // 3 秒內沒回覆的（4.6.0 以前的舊版 App 不會回）→ 照以前直接重新整理，才不會一直卡在舊版
      if (hadOld) {
        const wins = await self.clients.matchAll({ type: 'window' })
        const ok = await Promise.all(wins.map(askClient))
        wins.forEach((w, i) => ok[i] || w.navigate(w.url).catch(() => {}))
      }
    })(),
  )
})

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url)
  // Gemini API 不快取
  if (url.origin !== self.location.origin || e.request.method !== 'GET') return
  // 有網路：一律先問 GitHub 有沒有更新（no-cache＝沒變就很快回 304）；沒網路才用快取
  e.respondWith(
    caches.open(CACHE).then((cache) =>
      fetch(e.request, { cache: 'no-cache' })
        .then((res) => {
          if (res.ok) cache.put(e.request, res.clone())
          return res
        })
        .catch(() => cache.match(e.request, { ignoreSearch: true })),
    ),
  )
})
