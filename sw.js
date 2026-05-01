/* ============================================================
   sw.js — Service Worker for 3D QR Menu
   Cache-First strategy for all local assets
   ============================================================ */

const CACHE_NAME = '3dqrmenu-v2';

const ASSETS_TO_CACHE = [
  './',
  'index.html',
  'style.css',
  'app.js',
  'herovideo.mp4',
  'A_bold,_cinematic_202604202211.png',
  'Excel for my website.xlsx',
  'manifest.json'
];

// Install — precache core assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
  // Skip waiting so the new SW activates immediately
  self.skipWaiting();
});

// Activate — clean up old caches and claim clients
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(names => {
      return Promise.all(
        names
          .filter(name => name !== CACHE_NAME)
          .map(name => caches.delete(name))
      );
    })
  );
  // Claim all open clients immediately
  self.clients.claim();
});

// Fetch — Cache First for most assets, Network First for Excel data
self.addEventListener('fetch', event => {
  // Only handle GET requests
  if (event.request.method !== 'GET') return;

  // Network-first for Excel files (always get fresh data)
  if (event.request.url.endsWith('.xlsx')) {
    event.respondWith(
      fetch(event.request).then(networkResponse => {
        if (networkResponse && networkResponse.status === 200) {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then(cache => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      }).catch(() => {
        return caches.match(event.request);
      })
    );
    return;
  }

  // Cache-first for everything else
  event.respondWith(
    caches.match(event.request).then(cachedResponse => {
      if (cachedResponse) {
        return cachedResponse;
      }
      return fetch(event.request).then(networkResponse => {
        // Cache the new response for future use
        if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then(cache => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      }).catch(() => {
        // If both cache and network fail, return a fallback for navigation
        if (event.request.mode === 'navigate') {
          return caches.match('index.html');
        }
      });
    })
  );
});
