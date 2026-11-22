// TIMELINE 앱 껍데기용 서비스 워커
//  · 껍데기 파일(아이콘 등)만 저장해 두고, 일정 데이터(구글 웹앱)는 건드리지 않음
//  · index.html 은 항상 인터넷에서 새로 받음 → 주소를 고치면 바로 반영 (인터넷이 없을 때만 저장본)
const CACHE = "timeline-shell-v1";
const FILES = ["./", "index.html", "manifest.json", "icon-192.png", "icon-512.png", "icon-maskable-512.png", "apple-touch-icon.png", "favicon.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;   // 구글 웹앱 요청은 그대로 통과
  e.respondWith(
    fetch(e.request)
      .then((res) => { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); return res; })
      .catch(() => caches.match(e.request))
  );
});
