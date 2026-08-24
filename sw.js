/**
 * 台語寫字學習單編輯器 - Service Worker
 * 網路優先，失敗才用快取，方便更新；安裝後仍可離線開啟編輯器。
 */
var CACHE_NAME = "taigi-writing-v1.1.5";
var PRECACHE = [
  "./index.html",
  "./styles.css?v=1.1.5",
  "./app.js?v=1.1.5",
  "./manifest.webmanifest",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-512-maskable.png",
  "./icons/apple-touch-icon.png",
  "./icons/favicon-32.png",
  "./templates/index.json",
  "./templates/intro.json",
  "./textbooks/index.json",
];

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches
      .open(CACHE_NAME)
      .then(function (cache) {
        return Promise.all(
          PRECACHE.map(function (url) {
            return cache.add(url).catch(function () {
              /* 單檔失敗毋影響其他資源 */
            });
          }),
        );
      })
      .then(function () {
        return self.skipWaiting();
      }),
  );
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches
      .keys()
      .then(function (keys) {
        return Promise.all(
          keys.map(function (key) {
            if (key !== CACHE_NAME) return caches.delete(key);
          }),
        );
      })
      .then(function () {
        return self.clients.claim();
      }),
  );
});

self.addEventListener("fetch", function (event) {
  var request = event.request;
  if (request.method !== "GET") return;

  var url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(request)
      .then(function (response) {
        if (response && response.ok) {
          var copy = response.clone();
          caches.open(CACHE_NAME).then(function (cache) {
            cache.put(request, copy);
          });
        }
        return response;
      })
      .catch(function () {
        return caches.match(request).then(function (cached) {
          if (cached) return cached;
          if (request.mode === "navigate") {
            return caches.match("./index.html");
          }
          return caches.match(url.pathname).then(function (byPath) {
            return byPath || caches.match("." + url.pathname);
          });
        });
      }),
  );
});
