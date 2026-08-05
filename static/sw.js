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
var CACHE_NAME = 'gratz-shell-v1';
var RUNTIME_CACHE = 'gratz-pages-runtime';
var SHELL_FILES = [
  '/login',
  '/static/css/style.css',
  '/static/js/app.js',
  '/static/js/offline.js',
  '/static/js/vendor/chart.umd.min.js',
  '/static/js/vendor/feather.min.js',
  '/static/manifest.json',
  '/static/offline.html',
  '/static/img/icons/icon-192.png',
  '/static/img/icons/icon-512.png',
  '/static/img/icons/apple-touch-icon.png',
  '/static/img/logo-mark.png',
  '/static/img/logo-horizontal-dark.png',
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
          if (cached) return cached;
          // No exact match (same path + same query string) — try again
          // ignoring the query string. offline.js's precacheKeyPages()
          // warms each page's bare URL (e.g. /reports/income, no
          // ?period=...) in the background before anyone's ever visited
          // it; a real request with a filter/date-range query attached
          // would otherwise miss that entirely and fall through to the
          // generic offline page even though a close, useful version of
          // that same page is sitting right there in the cache.
          return caches.match(request, { ignoreSearch: true }).then(function (approxCached) {
            return approxCached || caches.match('/static/offline.html');
          });
        });
      })
    );
    return;
  }

  // Network-first, cache-fallback: a redeploy's new CSS/JS/icons should
  // reach a returning device the very next time it's online, instead of
  // being stuck behind whatever got cached under this CACHE_NAME the first
  // time (pure cache-first never re-checks the network for an asset it
  // already has, so it self-heals only when a human remembers to bump
  // CACHE_NAME — easy to forget on a deploy that doesn't touch sw.js).
  // Falls back to cache on any failure — a real network error/offline, or
  // an on-the-wire hiccup (a Render restart returning a transient 5xx
  // mid-deploy) — since serving a stale-but-working asset beats an
  // unstyled page. Only a non-ok response with nothing cached passes that
  // response through as-is (nothing better available); a network error
  // with nothing cached — first-ever load, offline, before install's
  // precache finished — has nothing to fall back to either.
  var url = new URL(request.url);
  if (request.method === 'GET' && url.pathname.startsWith('/static/')) {
    event.respondWith(
      fetch(request).then(function (response) {
        if (response.ok) {
          var copy = response.clone();
          caches.open(CACHE_NAME).then(function (cache) { cache.put(request, copy); });
          return response;
        }
        return caches.match(request).then(function (cached) { return cached || response; });
      }).catch(function () {
        return caches.match(request);
      })
    );
    return;
  }

  // Everything else (all form POSTs, /api/csrf-token, dashboard chart JSON)
  // passes through untouched — offline.js owns the queueing decision there.
});
