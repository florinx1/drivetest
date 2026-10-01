// ICG Drive Test – service worker
// Fișierele aplicației și fonturile se servesc imediat din memoria telefonului și se actualizează în fundal.
// Datele (Apps Script) NU trec prin cache: se citesc mereu de la server.
const CACHE = 'icg-drivetest-v8';
const SHELL = ['./', 'index.html', 'config.js', 'manifest.json', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/brand/foton.png', 'icons/brand/intercargo.png', 'icons/brand/foton-dark.png', 'icons/brand/intercargo-dark.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== CACHE).map(x => caches.delete(x)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request; if (req.method !== 'GET') return;
  const u = new URL(req.url);
  const sameOrigin = u.origin === location.origin;
  const staticCdn = /(^|\.)fonts\.(googleapis|gstatic)\.com$/.test(u.hostname) || u.hostname === 'cdnjs.cloudflare.com';
  if (!sameOrigin && !staticCdn) return;              // Apps Script, Drive: direct din rețea
  if (u.hostname === 'fonts.gstatic.com' || u.hostname === 'cdnjs.cloudflare.com') {   // fișiere cu versiune fixă: din cache
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(req, c)); return r; })));
    return;
  }
  // aplicația și CSS-ul fonturilor: imediat din cache, actualizare în fundal
  e.respondWith(caches.open(CACHE).then(cache => cache.match(req, { ignoreSearch: sameOrigin }).then(hit => {
    const net = fetch(req).then(r => { if (r && (r.ok || r.type === 'opaque')) cache.put(req, r.clone()); return r; }).catch(() => hit);
    return hit || net;
  })));
});
