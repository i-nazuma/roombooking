// Service Worker für die Raumbuchungs-PWA
//
// Zweck: NICHT Offline-Buchung ermöglichen (macht bei Live-Kalenderdaten
// keinen Sinn), sondern verhindern, dass ein kurzer WLAN-Aussetzer zu einer
// Browser-Fehlerseite auf dem Tablet führt. Bei Netzwerkfehlern wird die
// zuletzt erfolgreich geladene Ansicht aus dem Cache gezeigt.

const CACHE_NAME = "raumbuchung-shell-v1";

const APP_SHELL = [
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Wichtig: API-Calls (/api/...) NIEMALS aus dem Cache bedienen.
  // Kalenderstatus und Buchungen müssen immer live vom Backend kommen.
  if (url.pathname.startsWith("/api/")) {
    return;
  }

  // App-Shell: Network-first, mit Cache als Fallback bei Verbindungsproblemen.
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
