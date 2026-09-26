/* Service worker: guarda la app en el teléfono para que abra sin señal. */
const VERSION = 'metrados-v1.0.5';
const LIB = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.mini.min.js';
const FILES = ['./', './index.html', './style.css', './app.js', './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(async c => {
    await c.addAll(FILES.map(f => new Request(f, { cache: 'reload' })));
    try { await c.add(new Request(LIB, { mode: 'cors' })); } catch (err) { /* se guardará la primera vez que se use */ }
  }).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  if (url.origin !== location.origin && !url.href.startsWith('https://cdn.jsdelivr.net/npm/xlsx@')) return;
  // primero el caché (funciona sin señal); si hay red, se actualiza en segundo plano
  e.respondWith(caches.open(VERSION).then(async c => {
    const hit = await c.match(e.request, { ignoreSearch: true });
    const net = fetch(e.request, { cache: 'no-cache' }).then(r => { if (r && r.ok) c.put(e.request, r.clone()); return r; }).catch(() => null);
    if (hit) return hit;
    const r = await net;
    if (r) return r;
    return e.request.mode === 'navigate' ? c.match('./index.html') : Response.error();
  }));
});
