/* La app de Almacén se mudó a /almacen/: este service worker se elimina solo. */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => { e.waitUntil(caches.delete('almacen-v1.0.0').then(() => self.registration.unregister())); });
