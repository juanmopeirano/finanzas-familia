// Service worker — cache robusto que NUNCA devuelve null
const CACHE = 'finanzas-v12';
const SHELL = ['./','./index.html','./styles.css','./app.js','./icon.svg',
               './icon-192.png','./icon-512.png','./apple-touch-icon.png','./manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(SHELL))
    // OJO: no llamamos skipWaiting() automáticamente.
    // El nuevo SW espera hasta que el usuario apriete el banner "Actualizar".
  );
});

// Recibe mensaje del cliente para activar el SW nuevo en demanda
self.addEventListener('message', e => {
  if (e.data && e.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Solo manejamos GET de mismo origen — no interferimos con CDN ni APIs externas
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  if (url.origin !== self.location.origin) return;

  // Datos: stale-while-revalidate (sirve cache rápido, refresca en background)
  if (url.pathname.endsWith('finanzas.json')) {
    e.respondWith((async () => {
      const cache = await caches.open(CACHE);
      const cached = await cache.match(e.request, { ignoreSearch: true });

      if (cached) {
        // Refresca en background (sin bloquear la respuesta)
        fetch(e.request)
          .then(r => { if (r && r.ok) cache.put(e.request, r.clone()); })
          .catch(() => {});
        return cached;
      }

      // Sin cache: hay que ir a red, pero NUNCA devolver null
      try {
        const r = await fetch(e.request);
        if (r && r.ok) cache.put(e.request, r.clone());
        return r;
      } catch (err) {
        return new Response(
          JSON.stringify({ error: 'offline' }),
          { status: 503, headers: { 'Content-Type': 'application/json' } }
        );
      }
    })());
    return;
  }

  // Shell (HTML/CSS/JS/imgs): cache-first con fallback a red, y NUNCA null
  e.respondWith((async () => {
    const cached = await caches.match(e.request);
    if (cached) return cached;
    try {
      const r = await fetch(e.request);
      if (r && r.ok) {
        const cache = await caches.open(CACHE);
        cache.put(e.request, r.clone());
      }
      return r;
    } catch (err) {
      // Si es navegación, intentar servir el index del cache
      if (e.request.mode === 'navigate') {
        const fallback = await caches.match('./index.html');
        if (fallback) return fallback;
      }
      return new Response('Offline', { status: 503, headers: { 'Content-Type': 'text/plain' } });
    }
  })());
});
