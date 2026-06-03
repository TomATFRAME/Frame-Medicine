// FRAME Medicine — Patient App Service Worker
// Handles push notifications and basic offline caching.
// Strategy: network-first for HTML (so updates propagate on next open),
// cache-fallback only when offline. Other same-origin assets use cache-first.

var CACHE_NAME = 'frame-patient-v2';
var urlsToCache = ['/'];

self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(urlsToCache);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(cacheNames) {
      return Promise.all(
        cacheNames.filter(function(cacheName) {
          return cacheName !== CACHE_NAME;
        }).map(function(cacheName) {
          return caches.delete(cacheName);
        })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(event) {
  var req = event.request;
  if (req.method !== 'GET') return;

  var sameOrigin = (new URL(req.url)).origin === self.location.origin;
  if (!sameOrigin) return;

  var accept = req.headers.get('accept') || '';
  var isHTML = req.mode === 'navigate' || accept.indexOf('text/html') !== -1;

  if (isHTML) {
    event.respondWith(
      fetch(req).then(function(response) {
        var copy = response.clone();
        caches.open(CACHE_NAME).then(function(cache) { cache.put(req, copy); });
        return response;
      }).catch(function() {
        return caches.match(req).then(function(cached) {
          return cached || caches.match('/');
        });
      })
    );
    return;
  }

  event.respondWith(
    caches.match(req).then(function(cached) {
      return cached || fetch(req).then(function(response) {
        var copy = response.clone();
        caches.open(CACHE_NAME).then(function(cache) { cache.put(req, copy); });
        return response;
      });
    })
  );
});

self.addEventListener('push', function(event) {
  var data = { title: 'FRAME Medicine', body: 'You have a new notification' };
  if (event.data) {
    try {
      data = event.data.json();
    } catch (e) {
      data.body = event.data.text();
    }
  }
  event.waitUntil(
    self.registration.showNotification(data.title, {
      body: data.body,
      icon: 'https://framemedicine.com/wp-content/uploads/2025/08/Untitled-design.png',
      badge: 'https://framemedicine.com/wp-content/uploads/2025/08/Untitled-design.png',
      vibrate: [200, 100, 200],
      data: { url: 'https://framemedicine.com/app' }
    })
  );
});

self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  var url = event.notification.data ? event.notification.data.url : 'https://framemedicine.com/app';
  event.waitUntil(
    clients.matchAll({ type: 'window' }).then(function(clientList) {
      for (var i = 0; i < clientList.length; i++) {
        if (clientList[i].url.indexOf('/app') !== -1) {
          return clientList[i].focus();
        }
      }
      return clients.openWindow(url);
    })
  );
});
