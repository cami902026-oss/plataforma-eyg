// ===== ENERGY PWA Service Worker =====
// Estrategia: network-first para HTML/JS (siempre intenta traer la última versión),
// cache fallback cuando no hay internet → la app sigue abriendo del cache offline.

const CACHE_NAME = 'energy-v18';
const ASSETS = [
  './Index.html',
  './manifest.json',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS).catch(() => {}))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  // No interceptar llamadas al proxy de Claude ni a APIs externas
  const url = new URL(req.url);
  if (url.host.includes('script.google.com')) return;
  if (url.host.includes('api.github.com')) return;
  if (url.host.includes('graph.microsoft.com')) return;
  if (url.host.includes('login.microsoftonline.com')) return;
  if (url.host.includes('api.anthropic.com')) return;
  if (url.host.includes('netlify.app') && url.pathname.startsWith('/.netlify/')) return;

  event.respondWith(
    fetch(req)
      .then((res) => {
        if (res.ok && url.origin === location.origin) {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req).then((cached) => cached || caches.match('./Index.html')))
  );
});
