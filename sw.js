const CACHE_NAME = 'finanzas-pwa-v1';
const assets = ['./index.html', './manifest.json'];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(assets);
    })
  );
});

self.addEventListener('fetch', (e) => {
  e.respondWith(
    caches.match(e.request).then((res) => {
      return res || fetch(e.request);
    })
  );
}); //Es el "Service Worker" que almacena los archivos en la caché local de tu dispositivo.
