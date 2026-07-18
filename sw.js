const CACHE = 'pasuni-ocr-v1';
const SHELL = ['./index.html', './app.js', './codebook.js', './docx-editor.js', './manifest.json', './icon.svg'];

self.addEventListener('install', e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).catch(()=>{}));
  self.skipWaiting();
});
self.addEventListener('activate', e=>{
  e.waitUntil(
    caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k))))
  );
  self.clients.claim();
});
self.addEventListener('fetch', e=>{
  // Never cache API calls to generativelanguage.googleapis.com
  if (e.request.url.includes('generativelanguage.googleapis.com') || e.request.url.includes('unpkg.com')) return;
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).catch(()=>cached))
  );
});
