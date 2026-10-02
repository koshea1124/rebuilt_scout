// Offline support: keeps the app itself (page, code, fonts, Firebase SDK) on the phone
// so it opens in the stands with no signal. Scouting data is handled by Firestore's own
// offline cache, so this worker never touches database traffic.
// Bump VERSION whenever you change index.html, config.js or other app files.
const VERSION = "v2";
const CACHE = "scout-" + VERSION;
const SHELL = [
  "./",
  "./index.html",
  "./config.js",
  "./manifest.webmanifest",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/apple-touch-icon.png"
];
const CDN = [
  "https://www.gstatic.com/firebasejs/",
  "https://fonts.googleapis.com/",
  "https://fonts.gstatic.com/"
];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith("scout-") && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  const sameOrigin = url.origin === self.location.origin;
  const cdn = CDN.some(p => req.url.startsWith(p));
  if (!sameOrigin && !cdn) return; // leave Firebase Auth/Firestore traffic alone

  // Network first for our own files (so updates show up), cache when offline.
  // Cache first for versioned CDN files (Firebase SDK, fonts).
  if (sameOrigin) {
    e.respondWith(
      fetch(req).then(res => {
        if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
        return res;
      }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match("./index.html")))
    );
  } else {
    e.respondWith(
      caches.match(req).then(hit => hit || fetch(req).then(res => {
        if (res.ok || res.type === "opaque") { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
        return res;
      }))
    );
  }
});
