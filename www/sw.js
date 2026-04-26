const CACHE_NAME = 'buchitto-v2';
const ASSETS = [
  './',
  './index.html',
  './style.css',
  './script.js',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

// Install: cache core assets
self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Activate: clean old caches
self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch: Network-first, fallback to cache
self.addEventListener('fetch', (e) => {
  e.respondWith(
    fetch(e.request).then(response => {
      // Cache successful GET requests
      if (response.ok && e.request.method === 'GET') {
        const clone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(e.request, clone));
      }
      return response;
    }).catch(() => {
      // Offline fallback: use cache
      return caches.match(e.request).then(cached => {
        if (cached) return cached;
        if (e.request.destination === 'document') {
          return caches.match('./index.html');
        }
      });
    })
  );
});
