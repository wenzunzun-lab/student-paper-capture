const CACHE_NAME = "answer-capture-v7";
const APP_SHELL = [
  "./mobile-capture.html",
  "./mobile-capture.css?v=20261015-v6",
  "./pinyin-pro.js?v=3.29.3",
  "./mobile-capture.js?v=20261015-v6",
  "./mobile-manifest.webmanifest",
  "./vendor/lucide.min.js",
  "./vendor/xlsx.full.min.js"
];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE_NAME).then(async cache => {
    await cache.addAll(APP_SHELL);
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
