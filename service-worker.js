// ===== ENERGY PWA Service Worker =====
// Estrategia: network-first para HTML/JS (siempre intenta traer la última versión),
// cache fallback cuando no hay internet → la app sigue abriendo del cache offline.

const CACHE_NAME = 'energy-v37';
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

  // Para los archivos propios (HTML/JS de la app) pedir SIEMPRE fresco, sin caché del
  // navegador. Antes `fetch(req)` respetaba el caché HTTP de GitHub Pages (~10 min), así
  // que aunque se desplegara un arreglo, los equipos seguían viendo la versión vieja al
  // recargar. Con cache:'no-store' la recarga normal ya trae la última versión.
  const sameOrigin = url.origin === location.origin;
  event.respondWith(
    fetch(req, sameOrigin ? { cache: 'no-store' } : undefined)
      .then((res) => {
        if (res.ok && sameOrigin) {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req).then((cached) => cached || caches.match('./Index.html')))
  );
});
