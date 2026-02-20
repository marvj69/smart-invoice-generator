const APP_CACHE = "invoice-generator-v2-shell-v4";
const RUNTIME_CACHE = "invoice-generator-v2-runtime-v4";
const CACHE_PREFIX = "invoice-generator-v2-";
const OFFLINE_FALLBACK = "./index.html";

const APP_SHELL_FILES = [
  "./index.html",
  "./manifest.webmanifest",
  "./icons/icon-192.svg",
  "./icons/icon-512.svg",
  "./icons/icon-maskable.svg"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    (async () => {
      const cache = await caches.open(APP_CACHE);
      await cache.addAll(APP_SHELL_FILES);
      await self.skipWaiting();
    })()
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      const oldKeys = keys.filter(
        (key) =>
          key.startsWith(CACHE_PREFIX) &&
          key !== APP_CACHE &&
          key !== RUNTIME_CACHE
      );
      await Promise.all(oldKeys.map((key) => caches.delete(key)));
      await self.clients.claim();
    })()
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") {
    return;
  }

  if (event.request.mode === "navigate") {
    event.respondWith(networkFirst(event.request));
    return;
  }

  event.respondWith(cacheFirst(event.request));
});

self.addEventListener("message", (event) => {
  if (event.data && event.data.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});

async function networkFirst(request) {
  try {
    return await fetchAndCache(request);
  } catch (error) {
    const cachedResponse = await caches.match(request);
    if (cachedResponse) {
      return cachedResponse;
    }

    const fallback = await caches.match(OFFLINE_FALLBACK);
    if (fallback) {
      return fallback;
    }

    throw error;
  }
}

async function cacheFirst(request) {
  const cached = await caches.match(request);
  if (cached) {
    void refreshCacheInBackground(request);
    return cached;
  }

  return fetchAndCache(request);
}

async function fetchAndCache(request) {
  const response = await fetch(request);
  if (response.ok || response.type === "opaque") {
    const cache = await caches.open(RUNTIME_CACHE);
    await cache.put(request, response.clone());
  }
  return response;
}

async function refreshCacheInBackground(request) {
  try {
    await fetchAndCache(request);
  } catch (_error) {
    // Ignore background refresh failures and keep existing cached content.
  }
}
