/* Service worker: la app abre sin señal. Sube V al cambiar archivos. */
const V = "pesaje-v20";
const SHELL = ["./", "index.html", "config.js", "manifest.webmanifest", "img/icon-192.png", "img/icon-512.png", "img/logo.png"];
self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(V).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== V).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
function networkFirst(req) {
  return new Promise((resolve) => {
    let done = false;
    const fallback = () => caches.match(req, { ignoreSearch: true }).then((r) => r || caches.match("index.html"));
    const timer = setTimeout(() => { if (!done) { done = true; fallback().then(resolve); } }, 4000);
    fetch(req).then((res) => {
      const copy = res.clone(); caches.open(V).then((c) => c.put(req, copy));
      if (done) return;
      done = true; clearTimeout(timer); resolve(res);
    }).catch(() => { if (!done) { done = true; clearTimeout(timer); fallback().then(resolve); } });
  });
}
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;   // Apps Script no se intercepta
  e.respondWith(networkFirst(req));
});
