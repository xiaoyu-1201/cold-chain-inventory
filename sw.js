// 讓 App 本身離線也能打開（AI 辨識還是要網路）；有新版時背景更新，下次打開就是新的
const CACHE = 'inventory-v5'
const SHELL = ['./', 'index.html', 'app.js', 'styles.css', 'manifest.webmanifest', 'icon.svg', 'apple-touch-icon.png', 'icon-192.png', 'icon-512.png']

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()))
})

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  )
})

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url)
  // Gemini API 不快取
  if (url.origin !== self.location.origin || e.request.method !== 'GET') return
  // 有網路就拿最新版（順便更新快取）；沒網路才用快取
  e.respondWith(
    caches.open(CACHE).then((cache) =>
      fetch(e.request)
        .then((res) => {
          if (res.ok) cache.put(e.request, res.clone())
          return res
        })
        .catch(() => cache.match(e.request)),
    ),
  )
})
