const CACHE_NAME = "gagyebu-shell-v12";
const SHELL_FILES = [
  "index.html",
  "styles.css",
  "app.js",
  "config.js",
  "manifest.json"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  const isShellFile = url.origin === location.origin && SHELL_FILES.some((f) => url.pathname.endsWith(f));
  if (!isShellFile) return; // let auth/Graph requests go straight to the network, uncached

  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
