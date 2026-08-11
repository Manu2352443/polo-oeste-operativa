const CACHE = "polo-oeste-handheld-v1";
const STATIC_ASSETS = [
  "/static/handheld.css",
  "/static/feedback.css",
  "/static/feedback.js",
  "/static/polo-oeste.svg",
  "/static/handheld-manifest.webmanifest"
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(STATIC_ASSETS)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET" || !event.request.url.includes("/static/")) return;
  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
