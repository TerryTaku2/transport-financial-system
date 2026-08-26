/*
 * Offline queueing for field data-entry forms.
 * Any <form data-offline-form> is intercepted: submissions that get no
 * response (no signal, timeout, 5xx) are queued in IndexedDB and replayed
 * automatically once the device is back online. Forms without that
 * attribute are untouched and submit exactly as before.
 */
(function () {
  'use strict';

  var DB_NAME = 'transfleet-offline';
  var DB_VERSION = 2;
  var STORE = 'queue';
  var REFDATA_STORE = 'refdata';
  var SYNC_INTERVAL_MS = 30000;
  var FETCH_TIMEOUT_MS = 8000;
  // Same cache name sw.js's navigate handler reads from — writing into it
  // directly from here (via the Cache Storage API, not through the service
  // worker's fetch handler) is what lets a page get cached before anyone
  // has ever actually navigated to it. Much longer interval than
  // SYNC_INTERVAL_MS: this walks ~30 pages per run, so running it every
  // 30s would be needless load for something that only needs to be
  // reasonably fresh, not real-time.
  var PAGES_CACHE_NAME = 'gratz-pages-runtime';
  var PRECACHE_INTERVAL_MS = 10 * 60 * 1000;

  var syncing = false;

  function openDb() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          var store = db.createObjectStore(STORE, { keyPath: 'clientId' });
          store.createIndex('status', 'status');
          store.createIndex('createdAt', 'createdAt');
        }
        if (!db.objectStoreNames.contains(REFDATA_STORE)) {
          db.createObjectStore(REFDATA_STORE, { keyPath: 'key' });
        }
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }

  function saveRefData(data) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(REFDATA_STORE, 'readwrite');
        tx.objectStore(REFDATA_STORE).put({ key: 'all', data: data, fetchedAt: Date.now() });
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function getRefData() {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(REFDATA_STORE, 'readonly');
        var req = tx.objectStore(REFDATA_STORE).get('all');
        req.onsuccess = function () { resolve(req.result ? req.result.data : null); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function putItem(item) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, 'readwrite');
        tx.objectStore(STORE).put(item);
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function deleteItem(clientId) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, 'readwrite');
        tx.objectStore(STORE).delete(clientId);
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function getAll() {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, 'readonly');
        var req = tx.objectStore(STORE).getAll();
        req.onsuccess = function () { resolve(req.result || []); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function labelForForm(form) {
    var h1 = document.querySelector('h1');
    var title = (h1 && h1.textContent.trim()) || document.title;
    var parts = [];
    var vehicleSelect = form.querySelector('select[name="vehicle_id"]');
    if (vehicleSelect && vehicleSelect.selectedIndex > 0) {
      parts.push(vehicleSelect.options[vehicleSelect.selectedIndex].textContent.trim());
    }
    var dateField = form.querySelector('input[type="date"]');
    if (dateField && dateField.value) parts.push(dateField.value);
    return parts.length ? (title + ' — ' + parts.join(' — ')) : title;
  }

  function freshCsrfToken() {
    return fetch('/api/csrf-token', { credentials: 'same-origin' })
      .then(function (resp) {
        if (!resp.ok) throw new Error('csrf-token-fetch-failed');
        return resp.json();
      })
      .then(function (data) { return data.csrf_token; });
  }

  function fetchWithTimeout(url, options, timeoutMs) {
    var controller = new AbortController();
    var timer = setTimeout(function () { controller.abort(); }, timeoutMs);
    var merged = Object.assign({}, options, { signal: controller.signal });
    return fetch(url, merged).finally(function () { clearTimeout(timer); });
  }

  function notifyChanged() {
    document.dispatchEvent(new CustomEvent('offline-queue-changed'));
  }

  // Real connectivity, not just navigator.onLine (which only reflects the
  // network adapter, not whether the server is actually reachable) — set
  // from real request outcomes wherever we already have one, and otherwise
  // from a periodic /api/ping probe below.
  var connectivity = navigator.onLine ? 'online' : 'offline';

  function setConnectivity(state) {
    if (connectivity === state) return;
    connectivity = state;
    document.dispatchEvent(new CustomEvent('connectivity-changed'));
  }

  function checkConnectivity() {
    return fetchWithTimeout('/api/ping', { method: 'GET', credentials: 'same-origin', cache: 'no-store' }, FETCH_TIMEOUT_MS)
      .then(function () { setConnectivity('online'); })
      .catch(function () { setConnectivity('offline'); });
  }

  // One code path for both a live form submit and a later queue replay:
  // always fetch a fresh CSRF token first, since a token minted at
  // page-load can go stale long before a queued item gets replayed.
  function submitViaFetch(url, dataObject) {
    return freshCsrfToken()
      .catch(function () { return null; })
      .then(function (csrfToken) {
        if (csrfToken === null) {
          setConnectivity('offline');
          return { outcome: 'network-error' };
        }
        var formData = new FormData();
        Object.keys(dataObject).forEach(function (key) {
          formData.set(key, dataObject[key]);
        });
        formData.set('csrf_token', csrfToken);
        return fetchWithTimeout(url, {
          method: 'POST',
          body: formData,
          credentials: 'same-origin',
        }, FETCH_TIMEOUT_MS)
          .then(function (resp) { setConnectivity('online'); return { outcome: 'response', response: resp }; })
          .catch(function () { setConnectivity('offline'); return { outcome: 'network-error' }; });
      });
  }

  // Rebuilds every <select data-refdata="vehicles|drivers|parts|
  // franchise_vehicles|expense_categories"> on the page from real data,
  // instead of whatever was baked into the HTML the last time this page
  // was cached by the service worker. Preserves the current selection and
  // the placeholder ("Select vehicle" etc.) option if there is one.
  function repopulateSelects(data) {
    if (!data) return;
    document.querySelectorAll('select[data-refdata]').forEach(function (select) {
      var list = data[select.getAttribute('data-refdata')];
      if (!list) return;
      var currentValue = select.value;
      var placeholder = select.options.length && select.options[0].value === '' ? select.options[0] : null;
      select.innerHTML = '';
      if (placeholder) select.appendChild(placeholder);
      list.forEach(function (item) {
        var opt = document.createElement('option');
        opt.value = item.id;
        opt.textContent = item.label;
        if ('selling_price' in item) opt.dataset.price = item.selling_price;
        if ('quantity_on_hand' in item) opt.dataset.stock = item.quantity_on_hand;
        if ('cost_price' in item) opt.dataset.cost = item.cost_price;
        if ('unit' in item) opt.dataset.unit = item.unit;
        if (item.default_amount != null) opt.setAttribute('data-default-amount', item.default_amount);
        select.appendChild(opt);
      });
      if (currentValue && select.querySelector('option[value="' + currentValue + '"]')) {
        select.value = currentValue;
      }
      // Lets the searchable-select combobox in app.js (if this select was
      // enhanced) refresh its visible text to match the rebuilt options.
      select.dispatchEvent(new CustomEvent('refdata-repopulated'));
    });
  }

  // Only runs on pages that actually have one of the 9 offline-capable
  // forms — cheap no-op everywhere else. Online, this just re-confirms
  // what the server already rendered and refreshes the IndexedDB cache for
  // next time; offline, it's what lets a form served from the service
  // worker's stale page cache show real dropdown options instead of
  // whatever was on the page the last time it was cached.
  function refreshRefData() {
    return fetchWithTimeout('/api/refdata', { credentials: 'same-origin', cache: 'no-store' }, FETCH_TIMEOUT_MS)
      .then(function (resp) {
        if (!resp.ok) throw new Error('refdata fetch failed');
        return resp.json();
      })
      .then(function (data) {
        setConnectivity('online');
        return saveRefData(data).then(function () { repopulateSelects(data); });
      })
      .catch(function () {
        return getRefData().then(function (cached) {
          if (cached) repopulateSelects(cached);
        });
      });
  }

  // Runs on every page, not just ones with a form — this is what makes a
  // page work offline the FIRST time it's opened offline, not just the
  // second time onward (sw.js's own runtime cache only ever fills in
  // reactively, from pages actually visited while online). Silently
  // no-ops offline (nothing to fetch) or if caches storage isn't
  // available (very old browser) — this is a background nicety, never
  // something a page should wait on or surface an error for.
  function precacheKeyPages() {
    if (connectivity !== 'online' || !window.caches) return Promise.resolve();
    return fetchWithTimeout('/api/precache-urls', { credentials: 'same-origin', cache: 'no-store' }, FETCH_TIMEOUT_MS)
      .then(function (resp) { return resp.ok ? resp.json() : { urls: [] }; })
      .then(function (data) {
        var urls = data.urls || [];
        return caches.open(PAGES_CACHE_NAME).then(function (cache) {
          // Sequential, not Promise.all — this is a background low-priority
          // task walking ~30 pages; no reason to burst them all at once
          // against the same server a real user might be waiting on.
          return urls.reduce(function (chain, url) {
            return chain.then(function () {
              return fetchWithTimeout(url, { credentials: 'same-origin' }, FETCH_TIMEOUT_MS)
                .then(function (resp) { if (resp.ok) return cache.put(url, resp); })
                .catch(function () {}); // one page failing must never stop the rest
            });
          }, Promise.resolve());
        });
      })
      .catch(function () {});
  }

  function enqueueSubmission(form) {
    var formData = new FormData(form);
    var clientId = crypto.randomUUID();
    formData.set('_client_id', clientId);
    var item = {
      clientId: clientId,
      url: form.action || window.location.href,
      formData: Object.fromEntries(formData.entries()),
      label: labelForForm(form),
      createdAt: Date.now(),
      status: 'pending',
      lastError: null,
      attempts: 0,
    };
    return putItem(item).then(notifyChanged);
  }

  function showInlineNotice(form, message) {
    var notice = form.querySelector('.offline-inline-notice');
    if (!notice) {
      notice = document.createElement('div');
      notice.className = 'alert alert-warning offline-inline-notice';
      form.prepend(notice);
    }
    notice.textContent = message;
  }

  function attachFormInterception() {
    document.querySelectorAll('form[data-offline-form]').forEach(function (form) {
      form.addEventListener('submit', function (event) {
        event.preventDefault();
        var dataObject = Object.fromEntries(new FormData(form).entries());

        submitViaFetch(form.action || window.location.href, dataObject).then(function (result) {
          if (result.outcome === 'network-error') {
            return enqueueSubmission(form).then(function () {
              showInlineNotice(form, "Saved offline — will sync automatically once you're back online.");
              var firstField = form.querySelector('input, select, textarea');
              form.reset();
              if (firstField) firstField.focus();
            });
          }

          var resp = result.response;
          if (resp.ok || resp.redirected) {
            window.location.href = resp.url;
            return;
          }
          if (resp.status < 500) {
            // Genuine validation error — resubmit for real so the
            // server-rendered flash/error page shows exactly as it does today.
            form.submit();
            return;
          }
          // 5xx — treat like a connectivity problem and queue it.
          return enqueueSubmission(form).then(function () {
            showInlineNotice(form, "Saved offline — will sync automatically once you're back online.");
            form.reset();
          });
        });
      });
    });
  }

  function syncQueue() {
    if (syncing || !navigator.onLine) return Promise.resolve();
    syncing = true;

    return getAll()
      .then(function (items) {
        var queue = items
          .filter(function (i) { return i.status === 'pending' || i.status === 'syncing'; })
          .sort(function (a, b) { return a.createdAt - b.createdAt; });

        function next(i) {
          if (i >= queue.length) return Promise.resolve();
          var item = queue[i];
          item.status = 'syncing';
          return putItem(item).then(notifyChanged).then(function () {
            return submitViaFetch(item.url, item.formData);
          }).then(function (result) {
            if (result.outcome === 'network-error') {
              item.status = 'pending';
              return putItem(item).then(notifyChanged); // stop; don't burn through the rest offline
            }
            var resp = result.response;
            if (resp.ok || resp.redirected) {
              return deleteItem(item.clientId).then(notifyChanged).then(function () { return next(i + 1); });
            }
            if (resp.status === 401 || resp.status === 403) {
              item.status = 'auth-error';
              item.lastError = 'Sign in again to sync this entry.';
              return putItem(item).then(notifyChanged).then(function () { return next(i + 1); });
            }
            item.status = 'error';
            item.lastError = 'Server rejected this entry (HTTP ' + resp.status + ').';
            item.attempts = (item.attempts || 0) + 1;
            return putItem(item).then(notifyChanged).then(function () { return next(i + 1); });
          });
        }

        return next(0);
      })
      .finally(function () { syncing = false; });
  }

  function updateBadge() {
    var badge = document.getElementById('offline-sync-badge');
    if (!badge) return;
    getAll().then(function (items) {
      if (!items.length) {
        badge.style.display = 'none';
        return;
      }
      var pending = items.filter(function (i) { return i.status === 'pending'; }).length;
      var syncingCount = items.filter(function (i) { return i.status === 'syncing'; }).length;
      var errorCount = items.filter(function (i) { return i.status === 'error'; }).length;
      var authErrorCount = items.filter(function (i) { return i.status === 'auth-error'; }).length;

      badge.style.display = '';
      badge.style.cursor = 'pointer';
      badge.onclick = function () { window.location.href = '/static/offline.html'; };

      if (authErrorCount) {
        badge.className = 'badge badge-amber';
        badge.textContent = 'Sign in to sync';
      } else if (errorCount) {
        badge.className = 'badge badge-red';
        badge.textContent = errorCount + ' need attention';
      } else if (syncingCount) {
        badge.className = 'badge badge-blue';
        badge.textContent = 'Syncing ' + syncingCount + '…';
      } else {
        badge.className = 'badge badge-gray';
        badge.textContent = pending + ' pending sync';
      }
    });
  }

  function retryItem(clientId) {
    return getAll().then(function (items) {
      var item = items.filter(function (i) { return i.clientId === clientId; })[0];
      if (!item) return;
      item.status = 'pending';
      item.lastError = null;
      return putItem(item).then(notifyChanged).then(syncQueue);
    });
  }

  function discardItem(clientId) {
    return deleteItem(clientId).then(notifyChanged);
  }

  function updateConnectivityBadge() {
    var badge = document.getElementById('connectivity-badge');
    if (!badge) return;
    if (connectivity === 'online') {
      badge.className = 'badge badge-green';
      badge.textContent = 'Online';
    } else {
      badge.className = 'badge badge-gray';
      badge.textContent = 'Offline';
    }
  }

  document.addEventListener('offline-queue-changed', updateBadge);
  document.addEventListener('connectivity-changed', updateConnectivityBadge);

  // Regaining connectivity is exactly when a stale offline cache most
  // needs refreshing — whatever changed while this device was offline
  // (its own queued writes syncing back via syncQueue, or anyone else's)
  // should be reflected before the next time it goes offline again.
  window.addEventListener('online', function () { checkConnectivity(); syncQueue(); precacheKeyPages(); });
  window.addEventListener('offline', function () { checkConnectivity(); });
  document.addEventListener('visibilitychange', function () {
    if (!document.hidden) { checkConnectivity(); syncQueue(); }
  });
  setInterval(function () { checkConnectivity(); syncQueue(); }, SYNC_INTERVAL_MS);
  setInterval(precacheKeyPages, PRECACHE_INTERVAL_MS);

  document.addEventListener('DOMContentLoaded', function () {
    attachFormInterception();
    updateBadge();
    updateConnectivityBadge();
    checkConnectivity();
    syncQueue();
    if (document.querySelector('select[data-refdata]')) {
      refreshRefData();
    }
    precacheKeyPages();
  });

  // Exposed for static/offline.html, which lists/retries/discards queued
  // items directly — it isn't a Jinja page so it can't render server state.
  window.GratzOffline = {
    getAll: getAll,
    syncQueue: syncQueue,
    retryItem: retryItem,
    discardItem: discardItem,
  };
})();
