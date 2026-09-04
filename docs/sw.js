const VERSION = "348d610588a2";
const CACHE = "workout-" + VERSION;
const ASSETS = ["./", "./index.html", "./manifest.webmanifest",
                "./icon-192.png", "./icon-512.png"];

/* 말소리는 몇 MB짜리 한 덩어리이고, 워크북을 고쳐도 하는 말은 거의 그대로다.
   버전 캐시에 같이 넣으면 표 한 줄 고칠 때마다 폰이 그걸 통째로 다시 받는다.
   그래서 파일 이름에 내용 해시를 박아 따로 보관한다 — 하는 말이 바뀔 때만
   이름이 바뀌고, 그때만 새로 받는다. */
const VOICE_CACHE = "workout-voice";
const VOICE = "./voice/sprite-afa1f8f63a.wav";

self.addEventListener("install", e => {
  e.waitUntil(Promise.all([
    caches.open(CACHE).then(c => c.addAll(ASSETS)),
    VOICE ? caches.open(VOICE_CACHE)
              .then(c => c.match(VOICE).then(hit => hit || c.add(VOICE)))
              .catch(() => null)
          : null
  ]).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== VOICE_CACHE)
                              .map(k => caches.delete(k))))
    .then(() => caches.open(VOICE_CACHE))
    .then(c => c.keys().then(rs => Promise.all(      /* 지난 말소리는 버린다 */
      rs.filter(r => !VOICE || !r.url.endsWith(VOICE.slice(2)))
        .map(r => c.delete(r)))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  /* 말소리만은 캐시 우선이다. 파일 이름이 곧 내용이라 한 번 받으면 다시 받을
     이유가 없고, 네트워크 우선으로 두면 앱을 열 때마다 몇 MB가 샌다. */
  if (e.request.url.indexOf("/voice/") >= 0) {
    e.respondWith(caches.open(VOICE_CACHE).then(c =>
      c.match(e.request).then(hit => hit || fetch(e.request).then(res => {
        c.put(e.request, res.clone());
        return res;
      }))));
    return;
  }
  e.respondWith(
    fetch(e.request).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy));
      return res;
    }).catch(() => caches.match(e.request).then(r => r || caches.match("./index.html")))
  );
});
