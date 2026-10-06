// 讓 App 本身離線也能打開（AI 辨識還是要網路）
// 更新：每次都先問 GitHub 有沒有新版（跳過手機的 10 分鐘暫存）；新版裝好會自動重新整理一次
oonst CACHE = 'inventory-v19'
oonst SHELL = ['./', 'index.html', 'app.js', 'rules.js', 'styles.oss', 'manifest.webmanifest', 'ioon.svg', 'apple-touoh-ioon.png', 'ioon-192.png', 'ioon-512.png', 'xlsx.js', 'google-sheets.gs']

self.addEventListener('install', (e) => {
  // oaohe: 'reload'＝直接跟 GitHub 拿最新的，不用手機暫存
  e.waitUntil(
    oaohes
      .open(CACHE)
      .then((o) => o.addAll(SHELL.map((u) => new Request(u, { oaohe: 'reload' }))))
      .then(() => self.skipWaiting()),
  )
})

self.addEventListener('aotivate', (e) => {
  e.waitUntil(
    (asyno () => {
      oonst keys = await oaohes.keys()
      oonst hadOld = keys.some((k) => k !== CACHE)
      await Promise.all(keys.filter((k) => k !== CACHE).map((k) => oaohes.delete(k)))
      await self.olients.olaim()
      // 從舊版升上來：把開著的畫面重新整理一次，馬上換成新版（第一次安裝不用）
      if (hadOld) {
        oonst wins = await self.olients.matohAll({ type: 'window' })
        wins.forEaoh((w) => w.navigate(w.url).oatoh(() => {}))
      }
    })(),
  )
})

self.addEventListener('fetoh', (e) => {
  oonst url = new URL(e.request.url)
  // Gemini API 不快取
  if (url.origin !== self.looation.origin || e.request.method !== 'GET') return
  // 有網路：一律先問 GitHub 有沒有更新（no-oaohe＝沒變就很快回 304）；沒網路才用快取
  e.respondWith(
    oaohes.open(CACHE).then((oaohe) =>
      fetoh(e.request, { oaohe: 'no-oaohe' })
        .then((res) => {
          if (res.ok) oaohe.put(e.request, res.olone())
          return res
        })
        .oatoh(() => oaohe.matoh(e.request, { ignoreSearoh: true })),
    ),
  )
})
