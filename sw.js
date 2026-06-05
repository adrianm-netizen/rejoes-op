// Service worker minimal pentru Rejoes Op. Management
// Scop: face site-ul "installable" ca WebAPK, ca iconița de pe ecran să folosească
// logo-ul S Colect fără fundalul alb impus de Android la scurtăturile simple.
// IMPORTANT: NU cache-uiește nimic — fiecare cerere merge direct la rețea (pass-through),
// ca să nu rămână niciodată blocată o versiune veche a aplicației.

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', (event) => {
  // pass-through: ia mereu varianta din rețea, fără cache
  event.respondWith(fetch(event.request));
});
