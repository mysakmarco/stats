// Service worker скрінера: робить сторінку застосунком і показує push-сповіщення від бота.
const CACHE = "screener-v2";
const SHELL = ["./", "./index.html", "./manifest.webmanifest", "./icons/icon-192.png", "./icons/icon-512.png", "./icons/badge-96.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {}).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});

// сторінка: спершу мережа (завжди свіжа версія), без мережі — збережена копія.
// Біржі, GitHub і все чуже не чіпаємо.
self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET") return;
  const url = new URL(r.url);
  if (url.origin !== self.location.origin || !url.pathname.startsWith(new URL("./", self.registration.scope).pathname)) return;
  e.respondWith((async () => {
    try {
      const res = await fetch(r);
      if (res.ok && (r.mode === "navigate" || SHELL.some(p => url.pathname.endsWith(p.slice(1)) && p !== "./"))) {
        const c = await caches.open(CACHE);
        c.put(r.mode === "navigate" ? "./" : r, res.clone());
      }
      return res;
    } catch (err) {
      const hit = await caches.match(r.mode === "navigate" ? "./" : r, {ignoreSearch: true});
      if (hit) return hit;
      throw err;
    }
  })());
});

// push від бота: {title, body, tag, url}
self.addEventListener("push", e => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (err) { d = {title: "Скрінер", body: e.data ? e.data.text() : ""}; }
  const title = d.title || "Скрінер";
  e.waitUntil(self.registration.showNotification(title, {
    body: d.body || "",
    tag: d.tag || undefined,
    renotify: !!d.tag,
    icon: "icons/icon-192.png",
    badge: "icons/badge-96.png",
    vibrate: [200, 100, 200],
    timestamp: Date.now(),
    data: {url: d.url || "./"},
  }));
});

// тап по сповіщенню: відкрити застосунок (або перейти у вже відкритий)
self.addEventListener("notificationclick", e => {
  e.notification.close();
  const target = new URL((e.notification.data && e.notification.data.url) || "./", self.registration.scope).href;
  e.waitUntil((async () => {
    const all = await self.clients.matchAll({type: "window", includeUncontrolled: true});
    for (const c of all) {
      if (c.url.startsWith(self.registration.scope)) {
        try { await c.focus(); } catch (err) {}
        try { c.postMessage({type: "open-setups"}); } catch (err) {}
        return;
      }
    }
    await self.clients.openWindow(target);
  })());
});
