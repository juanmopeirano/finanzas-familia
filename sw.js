// Service worker mínimo — cache-first para shell, network-first para datos
const CACHE = 'finanzas-v6';
const SHELL = ['./','./index.html','./styles.css','./app.js','./icon.svg',
               './icon-192.png','./icon-512.png','./apple-touch-icon.png','./manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Datos: network-first (siempre lo más fresco)
  if (url.pathname.endsWith('finanzas.json')) {
    e.respondWith(
      fetch(e.request).then(r => {
        const copy = r.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy));
        return r;
      }).catch(() => caches.match(e.request))
    );
    return;
  }
  // Shell: cache-first
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
