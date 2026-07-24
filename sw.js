const CACHE_NAME = "gagyebu-shell-v16";

// 앱 껍데기 — 없으면 동작이 막히는 필수 파일
const CORE_FILES = [
  "index.html",
  "styles.css",
  "app.js",
  "config.js",
  "manifest.json"
];

// 있으면 좋지만 없어도 앱은 동작하는 파일 (캐릭터 이미지)
const OPTIONAL_FILES = [
  "icons/char-saved.png",
  "icons/char-deleted.png",
  "icons/char-edited.png",
  "icons/char-empty.png",
  "icons/char-chart.png",
  "icons/char-login.png",
  "icons/logo.png",
  "icons/icon-192.png",
  "icons/icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    // 핵심 파일은 반드시 캐시
    await cache.addAll(CORE_FILES);
    // 나머지는 하나씩 시도해서, 실패해도 설치를 막지 않습니다.
    await Promise.all(
      OPTIONAL_FILES.map((f) => cache.add(f).catch(() => {}))
    );
  })());
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);
  if (url.origin !== location.origin) return; // 인증·Graph 요청은 그대로 네트워크로

  const all = CORE_FILES.concat(OPTIONAL_FILES);
  const isShell = all.some((f) => url.pathname.endsWith(f));
  if (!isShell) return;

  // 캐시 우선, 없으면 네트워크에서 받아 캐시에 채웁니다.
  event.respondWith((async () => {
    const cached = await caches.match(req);
    if (cached) return cached;
    try {
      const res = await fetch(req);
      if (res && res.ok) {
        const cache = await caches.open(CACHE_NAME);
        cache.put(req, res.clone());
      }
      return res;
    } catch (e) {
      return cached || Response.error();
    }
  })());
});
