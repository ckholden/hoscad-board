/**
 * HOSCAD Field Service Worker
 * Caches field app shell for offline resilience, handles push notifications.
 */

const CACHE_NAME = 'hoscad-field-v67';
// Audio files intentionally excluded — browser Range requests return 206 which
// cache.addAll() rejects atomically, breaking the entire pre-cache install.
// Audio is served from network on demand and cached at runtime by the fetch handler.
// Paths are absolute (origin-relative) to avoid ambiguity — this SW is at root
// but scoped to /field/, so relative paths would resolve correctly from the SW
// script location but absolute paths are more explicit and future-proof.
const APP_SHELL = [
  '/field/',
  '/api.js',
  '/download.png',
  '/manifest-field.json',
  '/icons/favicon.ico',
  '/icons/icon-32.png',
  '/icons/icon-192.png',
  '/icons/icon-192-maskable.png',
  '/icons/icon-512.png',
  '/icons/apple-touch-icon.png'
];

// Install — cache app shell
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(APP_SHELL).catch(() => {});
    })
  );
  self.skipWaiting();
});

// Activate — clean old caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
      );
    })
  );
  self.clients.claim();
});

// Fetch — network first, fallback to cache
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  const url = event.request.url;

  // Skip non-http(s) requests (e.g. chrome-extension://) — Cache API rejects them
  if (!url.startsWith('http')) return;

  // Skip external APIs and CDNs — let the page handle these directly (avoids CORS issues)
  if (url.includes('supabase.co') || url.includes('script.google.com') || url.includes('googleapis') ||
      url.includes('nominatim.openstreetmap.org') || url.includes('tile.openstreetmap.org') ||
      url.includes('unpkg.com') || url.includes('cdn.sheetjs.com') ||
      url.includes('api.adsb.lol')) {
    return;
  }

  // Network-first with cache update for app shell
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Only cache full 200 responses — partial (206) and error responses must not be cached
        if (response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => caches.match(event.request).then(r => r || new Response('OFFLINE', { status: 503 })))
  );
});

// Push notifications
self.addEventListener('push', (event) => {
  let data = {};
  try {
    data = event.data ? event.data.json() : {};
  } catch (e) {
    try { data = { data: { title: event.data.text() } }; } catch (e2) {}
  }

  const payload = data.data || data.notification || data;
  const title = payload.title || 'HOSCAD Field Alert';
  const body = payload.body || 'Dispatch update received';
  const isUrgent = payload.urgent === 'true' || payload.urgent === true;
  const tag = payload.tag || ('hoscad-field-' + Date.now());

  event.waitUntil(
    self.registration.showNotification(title, {
      body: body,
      icon: '/download.png',
      badge: '/download.png',
      tag: tag,
      requireInteraction: isUrgent,
      vibrate: [300, 100, 300, 100, 300],
      data: payload
    })
  );
});

// Notification click — focus or open field app
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clients) => {
      for (const client of clients) {
        if (client.url) {
          return client.focus();
        }
      }
      return self.clients.openWindow('/field/');
    })
  );
});
