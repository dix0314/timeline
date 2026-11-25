// TIMELINE 앱 서비스 워커 — 아이콘을 누르면 바로 열리게
//  · 앱 화면 파일(index.html · 아이콘): 기기에 저장해 둔 걸 바로 보여 주고, 뒤에서 새 버전을 받아 둠
//    → GitHub 에서 index.html 을 고치면 "그다음 번" 실행부터 반영됨
//  · 글꼴(구글 폰트 · 에스코어드림): 한 번 받으면 계속 기기에 저장된 걸 씀
//  · 일정 데이터(script.google.com)는 저장하지 않고 항상 서버에서 받음
const SHELL = "timeline-shell-v3";
const FONTS = "timeline-fonts-v1";
const FILES = ["./", "index.html", "manifest.json", "icon-192.png", "icon-512.png", "icon-maskable-512.png", "apple-touch-icon.png", "favicon.png"];
const FONT_HOSTS = ["fonts.googleapis.com", "fonts.gstatic.com", "cdn.jsdelivr.net"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(SHELL).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== SHELL && k !== FONTS).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;                       // 서버 요청(POST)은 그대로 통과
  const url = new URL(req.url);

  // 글꼴: 저장본 먼저, 없으면 받아서 저장
  if (FONT_HOSTS.includes(url.hostname)) {
    e.respondWith(
      caches.open(FONTS).then((c) => c.match(req).then((hit) => hit || fetch(req).then((res) => {
        if (res.ok || res.type === "opaque") c.put(req, res.clone());
        return res;
      })))
    );
    return;
  }

  // 앱 파일: 저장본을 바로 주고, 뒤에서 새로 받아 저장 (인터넷이 없어도 열림)
  if (url.origin === location.origin) {
    e.respondWith(
      caches.open(SHELL).then((c) => c.match(req, { ignoreSearch: true }).then((hit) => {
        const fresh = fetch(req).then((res) => { if (res.ok) c.put(req, res.clone()); return res; }).catch(() => hit);
        if (hit) { e.waitUntil(fresh); return hit; }
        return fresh;
      }))
    );
  }
});
