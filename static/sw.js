// Paragon Agro Distribution Management System - Offline Service Worker (v15.0)
const CACHE_NAME = 'paragon-distribution-cache-v15';
const ASSETS_TO_CACHE = [
    '/static/css/style.css',
    '/static/manifest.json',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css',
    'https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap'
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            console.log('[ServiceWorker v15] Caching App Shell');
            return cache.addAll(ASSETS_TO_CACHE).catch((err) => console.log('SW cache partial error:', err));
        }).then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((keyList) => {
            return Promise.all(keyList.map((key) => {
                if (key !== CACHE_NAME) {
                    console.log('[ServiceWorker v15] Purging obsolete cache:', key);
                    return caches.delete(key);
                }
            }));
        }).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', (event) => {
    // 1. For API calls and Navigation/HTML, ALWAYS try Network FIRST
    if (event.request.url.includes('/api/') || event.request.mode === 'navigate' || event.request.destination === 'document') {
        event.respondWith(
            fetch(event.request).catch(() => {
                return caches.match(event.request).then(cached => {
                    if (cached) return cached;
                    if (event.request.url.includes('/api/')) {
                        return new Response(JSON.stringify({ 
                            offline: true, 
                            message: "You are currently offline. Actions will be saved locally and auto-synced when connection is restored." 
                        }), {
                            headers: { 'Content-Type': 'application/json' }
                        });
                    }
                    return caches.match('/');
                });
            })
        );
        return;
    }

    // 2. For static assets (CSS, images, fonts), Network first with fallback to Cache
    event.respondWith(
        fetch(event.request).then((fetchRes) => {
            if (fetchRes && fetchRes.status === 200) {
                const clone = fetchRes.clone();
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, clone);
                });
            }
            return fetchRes;
        }).catch(() => caches.match(event.request))
    );
});
