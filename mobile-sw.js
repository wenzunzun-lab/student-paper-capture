const CACHE_NAME = "answer-capture-v3";
const APP_SHELL = [
  "./mobile-capture.html",
  "./mobile-capture.css?v=20260905-v2",
  "./mobile-capture.js?v=20260905-v3",
  "./mobile-manifest.webmanifest"
];
const OPTIONAL_ASSETS = [
  "https://unpkg.com/lucide@latest",
  "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"
];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE_NAME).then(async cache => {
    await cache.addAll(APP_SHELL);
    await Promise.allSettled(OPTIONAL_ASSETS.map(url => cache.add(url)));
  }));
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key)))));
  self.clients.claim();
});

self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;
  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request).then(response => {
    const copy = response.clone();
    caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
    return response;
  })));
});
