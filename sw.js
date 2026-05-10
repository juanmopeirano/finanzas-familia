// KILL SWITCH — este SW se autoborra y se desregistra.
// Sirve para limpiar versiones viejas que quedaron stuckeadas en navegadores
// (ej. iPhones de Pili que tienen un SW viejo cacheado).
// Una vez que todos los clientes corrieron este SW, se puede borrar el archivo.

self.addEventListener('install', e => {
  // Activación inmediata sin esperar
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    // Borrar todas las cachés viejas
    const keys = await caches.keys();
    await Promise.all(keys.map(k => caches.delete(k)));
    // Desregistrar este SW
    await self.registration.unregister();
    // Forzar reload de todas las ventanas abiertas para volver a estado limpio
    const clientList = await self.clients.matchAll({ type: 'window' });
    clientList.forEach(c => c.navigate(c.url));
  })());
});

// NO interceptamos ningún fetch — todo va directo a la red
