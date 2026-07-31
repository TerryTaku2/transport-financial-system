/*
 * App-shell service worker. Only caches static assets — never rendered
 * Jinja pages, since those are permission-gated per user and embed a CSRF
 * token that would go stale in a cache. Offline queueing itself lives in
 * offline.js, not here; this worker only handles navigation fallback and
 * shell asset caching.
 */
var CACHE_NAME = 'transfleet-shell-v1';
var SHELL_FILES = [
  '/static/css/style.css',
  '/static/js/app.js',
  '/static/js/offline.js',
  '/static/manifest.json',
  '/static/offline.html',
  '/static/img/icons/icon-192.png',
  '/static/img/icons/icon-512.png',
  '/static/img/icons/apple-touch-icon.png',
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) { return cache.addAll(SHELL_FILES); })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (key) { return key !== CACHE_NAME; })
            .map(function (key) { return caches.delete(key); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function (event) {
  var request = event.request;

  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request).catch(function () {
        return caches.match('/static/offline.html');
      })
    );
    return;
  }

  var url = new URL(request.url);
  if (request.method === 'GET' && url.pathname.startsWith('/static/')) {
    event.respondWith(
      caches.match(request).then(function (cached) {
        return cached || fetch(request);
      })
    );
    return;
  }

  // Everything else (all form POSTs, /api/csrf-token, dashboard chart JSON)
  // passes through untouched — offline.js owns the queueing decision there.
});
