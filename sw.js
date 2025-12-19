// Simple offline cache (works on https or localhost)
// Bumpa versionssträngen när du lägger till/byter assets så att mobilen får nya filer.
const CACHE = "jesper-room-v9";
const ASSETS = [
  "./",
  "./index.html",
  "./style.css",
  "./main.js",
  "./sw.js",
  "./assets/julbild.jpg",
  "./assets/tavla.gif",
  "./assets/julsang.mp3",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE)
      .then((cache) => cache.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Cache-first for assets, network-first for navigation
self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Only handle same-origin
  if (url.origin !== location.origin) return;

  // HTML navigation: try network first (fresh), fallback to cache
  if (req.mode === "navigate") {
    event.respondWith(
      fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put("./index.html", copy)).catch(()=>{});
        return res;
      }).catch(() => caches.match("./index.html"))
    );
    return;
  }

  // Other requests: cache first
  event.respondWith(
    caches.match(req).then((cached) => cached || fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(req, copy)).catch(()=>{});
      return res;
    }))
  );
});
