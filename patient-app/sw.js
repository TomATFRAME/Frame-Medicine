// FRAME Medicine — Patient App Service Worker
// Handles push notifications and basic offline caching

var CACHE_NAME = 'frame-patient-v1';
var urlsToCache = ['/app'];

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
  event.respondWith(
    caches.match(event.request).then(function(response) {
      return response || fetch(event.request);
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
