const CACHE = 'finance-tracker-v1';
const CORE_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE).then(function(cache) {
      return cache.addAll(CORE_ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE; }).map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Handles the Android "Share to" flow: the OS sends a POST with the shared
// image as multipart/form-data straight to this URL (no server involved,
// GitHub Pages can't run one). Stash the file in a cache and redirect to the
// app itself, which picks it up via ?shared=1 and runs it through the same
// OCR pipeline as manually tapping "Scan receipt".
async function handleShareTarget(event) {
  try {
    const formData = await event.request.formData();
    const file = formData.get('receipt_image');
    if (file) {
      const cache = await caches.open('shared-receipt');
      await cache.put('/shared-receipt-image', new Response(file));
    }
  } catch (e) {
    // fall through to redirect regardless, so the user isn't stuck on a blank POST response
  }
  return Response.redirect('./?shared=1', 303);
}

self.addEventListener('fetch', function(event) {
  var shareUrl = new URL(event.request.url);
  if (event.request.method === 'POST' && shareUrl.pathname.endsWith('/share-target/')) {
    event.respondWith(handleShareTarget(event));
    return;
  }

  if (event.request.method !== 'GET') return;
  var url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return; // let cross-origin (fonts, Chart.js, rate API) hit the network directly

  // Network-first: always serve the latest version when online, so a
  // pushed update shows up immediately. Only fall back to the cached
  // copy when there's no connection (offline use).
  event.respondWith(
    fetch(event.request).then(function(response) {
      if (response && response.status === 200) {
        var copy = response.clone();
        caches.open(CACHE).then(function(cache) { cache.put(event.request, copy); });
      }
      return response;
    }).catch(function() {
      return caches.match(event.request);
    })
  );
});
