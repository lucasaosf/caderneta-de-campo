const CACHE = 'caderneta-v6';
const CORE = ['./', './index.html', './manifest.webmanifest', './especies.js?v=6', './planilha.js?v=6', './firebase-config.js?v=6',
  './icon-192.png', './icon-512.png', './icon-maskable-512.png', './icon-claro-192.png', './icon-claro-512.png', './favicon.png', './apple-touch-icon.png',
  './logo-creme.png', './logo-verde.png', './logo-planilha.png', './capa.jpg', './capa-clara.jpg', './privacidade.html'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === location.origin && url.pathname.endsWith('/version.json')) return;
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(r => { if (r.ok && (url.pathname.endsWith('/') || url.pathname.endsWith('/index.html'))) { const cp = r.clone(); caches.open(CACHE).then(c => c.put('./index.html', cp)); } return r; })
      .catch(() => caches.match(req).then(hit => hit || caches.match('./index.html'))));
    return;
  }
  const fonts = url.hostname.endsWith('fonts.googleapis.com') || url.hostname.endsWith('fonts.gstatic.com');
  const fbsdk = url.hostname === 'www.gstatic.com' && url.pathname.startsWith('/firebasejs/');
  if (url.origin === location.origin || fonts || fbsdk) {
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => {
      if (r && (r.ok || r.type === 'opaque')) { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); }
      return r;
    })));
  }
});
