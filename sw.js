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

// ESTRATEGIA OPTIMIZADA: Desactiva el almacenamiento en el navegador durante la mantención.
// Si el usuario desliza el dedo hacia abajo, la app puentea directo a GitHub Pages.
self.addEventListener('fetch', (event) => {
    event.respondWith(
        fetch(event.request, { cache: 'no-store' })
            .catch(() => caches.match(event.request))
    );
});
