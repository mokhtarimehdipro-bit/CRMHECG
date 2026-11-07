// Service worker de l'application « Prospects salon » : la page s'ouvre même sans réseau.
// Stratégie : réseau d'abord pour la page (toujours la dernière version si connecté),
// cache en secours ; les appels au CRM (Apps Script) ne sont jamais mis en cache.
const CACHE = 'hecg-salon-v2';
const SHELL = ['salon.html', 'auth-fetch.js?v=1', 'salon.webmanifest', 'salon-icon.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || req.url.indexOf('script.google.com') !== -1 || req.url.indexOf('googleusercontent.com') !== -1) return;
  e.respondWith(
    fetch(req).then(res => {
      if (res && (res.ok || res.type === 'opaque')) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then(hit => hit || caches.match('salon.html')))
  );
});
