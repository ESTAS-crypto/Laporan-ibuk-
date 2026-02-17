// ===== SERVICE WORKER - Offline Cache =====
const CACHE_NAME = 'laporan-kr-v11';

// Files to cache for offline use
const STATIC_ASSETS = [
    './',
    './index.html',
    './style.css',
    './script.js',
    './auth.js',
    './stylepdf.js',
    './voice-core.js',
    './footer.html',
    './manifest.json',
    './qris.jpg',
    './icons/icon-192.svg',
    './icons/icon-512.svg',
    './icons/icon-192.png',
    './icons/icon-512.png'
];

// CDN libraries to cache
const CDN_ASSETS = [
    'https://unpkg.com/docx@8.5.0/build/index.umd.js',
    'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'
];

// Install: cache all static assets
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => {
                console.log('[SW] Caching static assets...');
                // Cache static assets first
                return cache.addAll(STATIC_ASSETS)
                    .then(() => {
                        // Cache CDN assets (don't fail if any CDN is down)
                        return Promise.allSettled(
                            CDN_ASSETS.map(url =>
                                cache.add(url).catch(err =>
                                    console.warn(`[SW] Failed to cache CDN: ${url}`, err)
                                )
                            )
                        );
                    });
            })
            .then(() => {
                console.log('[SW] All assets cached!');
                return self.skipWaiting();
            })
    );
});

// Activate: clean old caches
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys()
            .then((cacheNames) => {
                return Promise.all(
                    cacheNames
                        .filter((name) => name !== CACHE_NAME)
                        .map((name) => {
                            console.log('[SW] Deleting old cache:', name);
                            return caches.delete(name);
                        })
                );
            })
            .then(() => self.clients.claim())
    );
});

// Fetch: Cache First — ONLY for our own files and known CDNs
// Fetch: Smart Strategy
self.addEventListener('fetch', (event) => {
    // Skip non-GET requests
    if (event.request.method !== 'GET') return;

    // Skip non-http requests
    if (!event.request.url.startsWith('http')) return;

    const url = new URL(event.request.url);

    // ONLY handle our own origin and known CDN domains
    const isOwnOrigin = url.origin === self.location.origin;
    const isKnownCDN = url.hostname === 'unpkg.com' ||
        url.hostname === 'cdnjs.cloudflare.com' ||
        url.hostname === 'cdn.jsdelivr.net';

    if (!isOwnOrigin && !isKnownCDN) return;

    // Network First strategy for HTML and JS files (vital for updates)
    if (url.pathname.endsWith('.html') || url.pathname.endsWith('.js') || url.pathname === '/') {
        event.respondWith(
            fetch(event.request)
                .then((networkResponse) => {
                    return caches.open(CACHE_NAME).then((cache) => {
                        cache.put(event.request, networkResponse.clone());
                        return networkResponse;
                    });
                })
                .catch(() => {
                    return caches.match(event.request);
                })
        );
        return;
    }

    // Cache First for other assets (images, fonts, libs)
    event.respondWith(
        caches.match(event.request)
            .then((cachedResponse) => {
                if (cachedResponse) return cachedResponse;
                return fetch(event.request).then((networkResponse) => {
                    return caches.open(CACHE_NAME).then((cache) => {
                        cache.put(event.request, networkResponse.clone());
                        return networkResponse;
                    });
                });
            })
    );
});
