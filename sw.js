self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(cacheNames.map((cache) => caches.delete(cache)));
        }).then(() => self.clients.claim())
    );
});

// ESTRATEGIA OPTIMIZADA: Fuerza la petición a internet ignorando la caché del navegador.
// Esto permite que el gesto manual de "arrastrar con el dedo" funcione de inmediato.
self.addEventListener('fetch', (event) => {
    event.respondWith(
        fetch(event.request, { cache: 'no-store' })
            .catch(() => caches.match(event.request))
    );
});
