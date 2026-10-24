/* Service Worker — offline support for the Fukuoka 2026 travel handbook.
 *
 * Strategy: stale-while-revalidate. A cached copy is served instantly (no
 * waiting on a weak mountain / subway signal) while a fresh copy is fetched in
 * the background. When the background fetch brings a changed
 * itinerary_data.json, open pages get a 'data-updated' message and re-render,
 * so edits to itinerary.xlsx still show up within seconds of going online.
 * With nothing cached yet, it falls back to the network.
 * Bump CACHE_VERSION when the precache list changes.
 */
const CACHE_VERSION = 'v2';
const CACHE_NAME = `japan2026-${CACHE_VERSION}`;
const PRECACHE_URLS = ['./index.html', './itinerary_data.json', './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png', './icon-maskable-512.png'];
const DATA_PATH = 'itinerary_data.json';

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(PRECACHE_URLS.map((u) => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k.startsWith('japan2026-') && k !== CACHE_NAME).map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

async function notifyIfDataChanged(cached, fresh) {
  if (!cached) return;
  const [before, after] = await Promise.all([cached.text(), fresh.text()]);
  if (before === after) return;
  const clients = await self.clients.matchAll({ type: 'window' });
  clients.forEach((client) => client.postMessage({ type: 'data-updated' }));
}

async function staleWhileRevalidate(event) {
  const { request } = event;
  const cache = await caches.open(CACHE_NAME);
  // Every page visit (/, /index.html, ?now=...) shares one app-shell entry
  const key = request.mode === 'navigate' ? './index.html' : request;
  const cached = await cache.match(key, { ignoreSearch: true });

  const refresh = fetch(request, { cache: 'no-cache' }).then(async (response) => {
    if (response && response.ok && response.type === 'basic') {
      const isData = new URL(request.url).pathname.endsWith(DATA_PATH);
      // Re-read from the cache: the copy already handed to the page has its body consumed
      const previous = isData ? await cache.match(key, { ignoreSearch: true }) : null;
      await cache.put(key, response.clone());
      if (previous) await notifyIfDataChanged(previous, response.clone());
    }
    return response;
  });

  if (cached) {
    event.waitUntil(refresh.catch(() => {}));
    return cached;
  }
  return refresh;
}

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return; // let external links (maps, coupons) pass through
  event.respondWith(staleWhileRevalidate(event));
});
