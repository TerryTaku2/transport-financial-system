/*
 * App-shell service worker. The shell cache (CSS/JS/icons/manifest/login
 * page) is precached at install time. Every other page is added to a
 * separate runtime cache the first time it's successfully loaded online, so
 * a device that goes offline — even on a cold app launch — can still reach
 * any page it's visited before, not just the static shell.
 *
 * Caching rendered Jinja pages was originally avoided over CSRF-staleness
 * concerns, but that's moot: offline.js always fetches a fresh CSRF token
 * and overwrites the form's token before every submit (see submitViaFetch
 * in offline.js), so a stale token embedded in a cached page is never used.
 * The real tradeoff is a shared/kiosk device could see a previous user's
 * cached page while offline — acceptable here since each field user installs
 * their own copy of the app on their own device.
 */
var CACHE_NAME = 'transfleet-shell-v2';
var RUNTIME_CACHE = 'transfleet-pages-runtime';
var SHELL_FILES = [
  '/login',
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
        keys.filter(function (key) { return key !== CACHE_NAME && key !== RUNTIME_CACHE; })
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
      fetch(request).then(function (response) {
        if (response.ok && request.method === 'GET') {
          var copy = response.clone();
          caches.open(RUNTIME_CACHE).then(function (cache) { cache.put(request, copy); });
        }
        return response;
      }).catch(function () {
        return caches.match(request).then(function (cached) {
          return cached || caches.match('/static/offline.html');
        });
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
