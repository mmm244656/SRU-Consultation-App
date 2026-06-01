const CACHE = 'sru-hub-v1';
const ASSETS = [
  '/SRU-Consultation-App/',
  '/SRU-Consultation-App/index.html',
  '/SRU-Consultation-App/manifest.json',
  '/SRU-Consultation-App/icon.png',
  '/SRU-Consultation-App/icon-192.png',
  '/SRU-Consultation-App/icon-512.png',
  '/SRU-Consultation-App/icon-ios.png'
];

self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k){ return k !== CACHE; }).map(function(k){ return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e) {
  // Network first for API calls
  if(e.request.url.includes('script.google.com')) {
    return;
  }
  e.respondWith(
    fetch(e.request).catch(function() {
      return caches.match(e.request);
    })
  );
});
