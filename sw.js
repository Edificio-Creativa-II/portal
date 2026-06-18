const CACHE_NAME = 'siges-cache-v1';

// Fuerza al Service Worker a activarse de inmediato sin esperar
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

// Toma el control de la aplicación inmediatamente cuando hay cambios en el servidor
self.addEventListener('activate', (event) => {
    event.waitUntil(self.clients.claim());
});

// ESTRATEGIA: Network First (Prioriza internet para ver los cambios al tiro)
self.addEventListener('fetch', (event) => {
    event.respondWith(
        fetch(event.request)
            .then((response) => {
                // Si la red responde bien, guardamos copia fresca en caché y la entregamos
                if (response && response.status === 200) {
                    const responseClone = response.clone();
                    caches.open(CACHE_NAME).then((cache) => {
                        cache.put(event.request, responseClone);
                    });
                }
                return response;
            })
            .catch(() => {
                // Si el vecino está sin señal en el ascensor, carga lo último guardado
                return caches.match(event.request);
            })
    );
});
