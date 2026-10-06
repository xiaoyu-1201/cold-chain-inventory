// 讓 App 本身離線也能打開（AI 辨識還是要網路）
// 更新：每次都先問 GitHub 有沒有新版（跳過手機的 10 分鐘暫存）；新版裝好會自動重新整理一次
const CACHE = 'inventory-v25'
const SHELL = ['./', 'index.html', 'app.js', 'rules.js', 'styles.css', 'manifest.webmanifest', 'icon.svg', 'apple-touch-icon.png', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'logo.svg', 'xlsx.js', 'google-sheets.gs']

self.addEventListener('install', (e) => {
  // cache: 'reload'＝直接跟 GitHub 拿最新的，不用手機暫存
  e.waitUntil(
    caches
      .open(CACHE)
      .then((c) => c.addAll(SHELL.map((u) => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting()),
  )
})

self.addEventListener('activate', (e) => {
  e.waitUntil(
    (async () => {
      const keys = await caches.keys()
      const hadOld = keys.some((k) => k !== CACHE)
      await Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
      await self.clients.claim()
      // 從舊版升上來：把開著的畫面重新整理一次，馬上換成新版（第一次安裝不用）
      if (hadOld) {
        const wins = await self.clients.matchAll({ type: 'window' })
        wins.forEach((w) => w.navigate(w.url).catch(() => {}))
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
