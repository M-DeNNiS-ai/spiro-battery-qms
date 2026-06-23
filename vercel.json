// ═══════════════════════════════════════════════════════════════════════════════
// SPIRO BATTERY QMS — SERVICE WORKER v1.0
// Handles: offline caching, background sync, install prompt
// ═══════════════════════════════════════════════════════════════════════════════

const CACHE_NAME     = "spiro-qms-v1";
const SYNC_TAG       = "spiro-onedrive-sync";
const PENDING_KEY    = "spiro_pending_sync";

// All assets to pre-cache on install — app works 100% offline after first load
const PRECACHE_URLS = [
  "/",
  "/index.html",
  "/manifest.json",
  "/icons/icon-192x192.png",
  "/icons/icon-512x512.png",
];

// CDN scripts we cache on first use (MSAL, jsQR, React, etc.)
const CDN_CACHE_NAME = "spiro-cdn-v1";

// ── Install: pre-cache app shell ──────────────────────────────────────────────
self.addEventListener("install", (event) => {
  console.log("[SW] Installing Spiro QMS Service Worker…");
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log("[SW] Pre-caching app shell");
      return cache.addAll(PRECACHE_URLS);
    }).then(() => self.skipWaiting())
  );
});

// ── Activate: clean up old caches ────────────────────────────────────────────
self.addEventListener("activate", (event) => {
  console.log("[SW] Activating…");
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME && k !== CDN_CACHE_NAME)
            .map(k => { console.log("[SW] Deleting old cache:", k); return caches.delete(k); })
      )
    ).then(() => self.clients.claim())
  );
});

// ── Fetch: serve from cache, fall back to network ────────────────────────────
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Skip non-GET and Microsoft Graph API calls (always need network)
  if (event.request.method !== "GET") return;
  if (url.hostname.includes("graph.microsoft.com")) return;
  if (url.hostname.includes("login.microsoftonline.com")) return;
  if (url.hostname.includes("login.microsoft.com")) return;

  // CDN resources (React, MSAL, jsQR) — cache on first fetch, serve from cache after
  if (url.hostname.includes("cdnjs.cloudflare.com") || url.hostname.includes("unpkg.com") || url.hostname.includes("esm.sh")) {
    event.respondWith(
      caches.open(CDN_CACHE_NAME).then(async (cache) => {
        const cached = await cache.match(event.request);
        if (cached) return cached;
        try {
          const response = await fetch(event.request);
          if (response.ok) cache.put(event.request, response.clone());
          return response;
        } catch {
          return new Response("// CDN resource unavailable offline", { headers: { "Content-Type": "application/javascript" } });
        }
      })
    );
    return;
  }

  // App shell — cache first, network fallback
  event.respondWith(
    caches.match(event.request).then(async (cached) => {
      if (cached) return cached;
      try {
        const response = await fetch(event.request);
        if (response.ok) {
          const cache = await caches.open(CACHE_NAME);
          cache.put(event.request, response.clone());
        }
        return response;
      } catch {
        // Offline fallback — serve index.html for navigation requests
        if (event.request.mode === "navigate") {
          return caches.match("/index.html");
        }
        return new Response("Offline", { status: 503 });
      }
    })
  );
});

// ── Background Sync: retry failed OneDrive uploads ───────────────────────────
self.addEventListener("sync", (event) => {
  if (event.tag === SYNC_TAG) {
    console.log("[SW] Background sync triggered — retrying OneDrive uploads");
    event.waitUntil(retrySyncQueue());
  }
});

async function retrySyncQueue() {
  // Notify the app to retry — the app holds the queue in localStorage
  const clients = await self.clients.matchAll({ type: "window" });
  clients.forEach((client) => {
    client.postMessage({ type: "RETRY_SYNC" });
  });
}

// ── Push Notifications (future use) ──────────────────────────────────────────
self.addEventListener("push", (event) => {
  const data = event.data?.json() || {};
  event.waitUntil(
    self.registration.showNotification(data.title || "Spiro Battery QMS", {
      body: data.body || "You have a new notification.",
      icon: "/icons/icon-192x192.png",
      badge: "/icons/icon-72x72.png",
      vibrate: [100, 50, 100],
      data: { url: data.url || "/" },
    })
  );
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  event.waitUntil(
    clients.openWindow(event.notification.data?.url || "/")
  );
});

// ── Message handler: receive commands from the app ───────────────────────────
self.addEventListener("message", (event) => {
  if (event.data?.type === "SKIP_WAITING") self.skipWaiting();
  if (event.data?.type === "CACHE_URLS") {
    caches.open(CACHE_NAME).then(c => c.addAll(event.data.urls || []));
  }
});

console.log("[SW] Spiro QMS Service Worker loaded ✓");
