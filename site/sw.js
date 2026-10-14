/* Basaer service worker: pages network-first (fresh when online, cached when not),
 * assets and data cache-first. Bump VERSION to drop old caches after a deploy. */
const VERSION = "basaer-v1";
const CORE = [
  "./", "index.html", "tajweed.html", "surah.html", "sources.html", "install.html",
  "assets/basaer.css", "assets/basaer.js", "assets/fonts/UthmanicHafs_V22.woff2",
  "data/surahs.json", "data/tajweed-rules.json", "data/stats.json", "data/sources.json",
  "manifest.webmanifest", "icons/icon-192.png"
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(CORE)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET" || new URL(req.url).origin !== location.origin) return;
  if (req.mode === "navigate") {
    e.respondWith(fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(VERSION).then((c) => c.put(req, copy));
      return res;
    }).catch(() => caches.match(req).then((r) => r || caches.match("index.html"))));
    return;
  }
  e.respondWith(caches.match(req).then((hit) => hit || fetch(req).then((res) => {
    if (res.ok) { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(req, copy)); }
    return res;
  })));
});
