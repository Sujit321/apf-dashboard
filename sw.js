const CACHE_NAME = 'apf-dashboard-v2';
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './app.js',
    './styles.css',
    './excel-analytics.js',
    './icon-192.png',
    './icon-512.png'
];

self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => {
            console.log('Opened cache');
            return cache.addAll(ASSETS_TO_CACHE);
        })
    );
    self.skipWaiting();
});

self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(cacheNames => {
            return Promise.all(
                cacheNames.map(cache => {
                    if (cache !== CACHE_NAME) {
                        return caches.delete(cache);
                    }
                })
            );
        })
    );
    self.clients.claim();
});

self.addEventListener('fetch', event => {
    // Only cache GET requests
    if (event.request.method !== 'GET') return;

    event.respondWith(
        caches.match(event.request).then(response => {
            // Return from cache if found
            if (response) {
                return response;
            }

            // Else fetch from network
            return fetch(event.request).then(fetchRes => {
                // Check if valid response
                if (!fetchRes || fetchRes.status !== 200 || fetchRes.type !== 'basic') {
                    return fetchRes;
                }

                // Clone response to put in cache
                const responseToCache = fetchRes.clone();

                caches.open(CACHE_NAME).then(cache => {
                    // Don't cache chrome-extension requests or data URIs
                    if (event.request.url.startsWith('http')) {
                        cache.put(event.request, responseToCache);
                    }
                });

                return fetchRes;
            }).catch(() => {
                // If network fails and we don't have it in cache, we could return a fallback here
            });
        })
    );
});

