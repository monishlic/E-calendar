/* ============================================================================
   E-Calendar — Service Worker
   Makes the app installable and loads the interface instantly / offline.
   It caches ONLY the app shell (the HTML + icons). It NEVER caches your
   Google Apps Script API calls, so your leave data is always live.

   >>> To push an update to everyone after you change index.html on GitHub:
   >>> bump the version number below (ecal-v1 -> ecal-v2). That's it.
   ============================================================================ */
const CACHE = 'ecal-v3';

const SHELL = [
  './',
  './index.html',
  './manifest.json',
  './favicon.png',
  './icon-192.png',
  './icon-512.png',
  './icon-512-maskable.png',
  './apple-touch-icon.png'
];

// Pre-cache the shell on install (tolerant: a single missing file won't break it)
self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    await Promise.allSettled(SHELL.map((url) => cache.add(url)));
    await self.skipWaiting();
  })());
});

// Clean up old versions when a new SW activates
self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Only handle GETs. POSTs (writes to your API) always go straight to network.
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Only handle same-origin requests. Cross-origin (your Apps Script API at
  // script.google.com, Google Fonts, Gemini) always go straight to network —
  // so data is never stale and never cached.
  if (url.origin !== self.location.origin) return;

  // HTML page loads: network-first (always get the latest when online),
  // fall back to the cached shell when offline.
  if (req.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        const fresh = await fetch(req);
        const cache = await caches.open(CACHE);
        cache.put('./index.html', fresh.clone());
        return fresh;
      } catch (e) {
        const cached = await caches.match('./index.html');
        return cached || Response.error();
      }
    })());
    return;
  }

  // Static assets (icons, manifest): cache-first, then network.
  event.respondWith((async () => {
    const cached = await caches.match(req);
    if (cached) return cached;
    try {
      const fresh = await fetch(req);
      const cache = await caches.open(CACHE);
      cache.put(req, fresh.clone());
      return fresh;
    } catch (e) {
      return cached || Response.error();
    }
  })());
});
