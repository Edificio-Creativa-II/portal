const CACHE_NAME = 'siges-cache-202606181623';

// Fuerza al Service Worker a activarse de inmediato sin esperar
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

// Limpia las cachés viejas para que no ocupen espacio en el celular del vecino
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.map((cache) => {
                    if (cache !== CACHE_NAME) {
                        console.log('Borrando caché antigua:', cache);
                        return caches.delete(cache);
                    }
                })
            );
        }).then(() => self.clients.claim())
    );
});

// ESTRATEGIA: Network First (Prioriza internet para ver los cambios al tiro)
self.addEventListener('fetch', (event) => {
    event.respondWith(
        fetch(event.request)
            .then((response) => {
                if (response && response.status === 200) {
                    const responseClone = response.clone();
                    caches.open(CACHE_NAME).then((cache) => {
                        cache.put(event.request, responseClone);
                    });
                }
                return response;
            })
            .catch(() => {
                return caches.match(event.request);
            })
    );
});
