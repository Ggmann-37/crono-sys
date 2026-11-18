// Service Worker para Crono-SYS (Compatible con PWABuilder)
// Estrategia: Network-first con fallback a Cache y página offline

const CACHE_NAME = 'crono-sys-pwa-v1';
const OFFLINE_URL = 'offline.html';

const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './offline.html',
  './manifest.json',
  './css/styles.css',
  './js/app.js',
  './js/calendar-view.js',
  './js/excel-parser.js',
  './js/scheduler.js',
  './js/supabase-client.js',
  './assets/logo.svg',
  './assets/favicon.svg',
  './assets/icon-192.png',
  './assets/icon-512.png',
  './plantilla_ejemplo.csv'
];

// Instalación del Service Worker y precaching de assets esenciales
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[Service Worker] Precaching recursos esenciales');
      return cache.addAll(ASSETS_TO_CACHE).catch((err) => {
        console.warn('[Service Worker] Algunos recursos no pudieron precachearse:', err);
      });
    })
  );
  self.skipWaiting();
});

// Activación y limpieza de caches antiguas
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keyList) => {
      return Promise.all(
        keyList.map((key) => {
          if (key !== CACHE_NAME) {
            console.log('[Service Worker] Eliminando cache antigua:', key);
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// Estrategia de Fetch: Network-First con fallback a Cache y modo Offline
self.addEventListener('fetch', (event) => {
  // Solo manejar solicitudes GET
  if (event.request.method !== 'GET') return;

  // Evitar interceptar llamadas a Supabase Auth o API externa si se prefiere paso directo
  const url = new URL(event.request.url);
  if (url.origin.includes('supabase.co')) {
    return;
  }

  // Para navegaciones HTML (páginas)
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request)
        .then((networkResponse) => {
          return caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, networkResponse.clone());
            return networkResponse;
          });
        })
        .catch(() => {
          return caches.match(event.request).then((cachedResponse) => {
            if (cachedResponse) {
              return cachedResponse;
            }
            return caches.match(OFFLINE_URL);
          });
        })
    );
    return;
  }

  // Para otros assets (css, js, imágenes)
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        // Devolver de cache y actualizar en segundo plano (Stale-While-Revalidate)
        fetch(event.request)
          .then((networkResponse) => {
            if (networkResponse && networkResponse.status === 200) {
              caches.open(CACHE_NAME).then((cache) => {
                cache.put(event.request, networkResponse);
              });
            }
          })
          .catch(() => {});
        return cachedResponse;
      }

      // Si no está en cache, buscar en red
      return fetch(event.request)
        .then((networkResponse) => {
          if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic') {
            return networkResponse;
          }
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
          return networkResponse;
        })
        .catch(() => {
          // Si falla y es imagen, podríamos devolver logo svg
          if (event.request.headers.get('accept')?.includes('image')) {
            return caches.match('./assets/logo.svg');
          }
        });
    })
  );
});
