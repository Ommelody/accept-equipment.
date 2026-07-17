/* Service Worker — ASSET System PWA
   - HTML shell: network-first (ข้อมูล/URL ที่ฝังไว้ต้องเป็นเวอร์ชันล่าสุดเสมอเมื่อออนไลน์)
   - Static assets (ไอคอน/manifest): cache-first
   - ไม่แคชคำขอไปยัง backend (GAS) เพื่อให้ข้อมูลสดเสมอ
   - ออฟไลน์: ใช้แคชล่าสุดที่มี */
const CACHE = 'asset-shell-v2';
const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;

  // ไม่ยุ่งกับคำขอที่ไม่ใช่ GET (POST ไป backend ฯลฯ)
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // คำขอไปยัง backend (Google Apps Script) — network only, ไม่แคช
  if (/script\.google\.com|googleusercontent\.com/.test(url.hostname)) {
    return; // ปล่อยให้เบราว์เซอร์จัดการปกติ
  }

  const isHtmlShell = url.origin === location.origin && (req.mode === 'navigate' || req.destination === 'document' || url.pathname.endsWith('/') || url.pathname.endsWith('index.html'));

  if (isHtmlShell) {
    // network-first: ต้องได้หน้าล่าสุดทุกครั้งที่ออนไลน์ (กัน URL/ค่าฝังตัวค้างจากแคชเก่า)
    e.respondWith(
      fetch(req, { cache: 'no-store' }).then((res) => {
        if (res && res.status === 200) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
        return res;
      }).catch(() => caches.match(req))
    );
    return;
  }

  // asset อื่น ๆ (ไอคอน/manifest/CDN): cache-first + อัปเดตพื้นหลัง
  e.respondWith(
    caches.match(req).then((cached) => {
      const fetchPromise = fetch(req).then((res) => {
        if (res && res.status === 200 && (url.origin === location.origin || /jsdelivr|cloudflare|googleapis|gstatic|jquery/.test(url.hostname))) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
        }
        return res;
      }).catch(() => cached);
      return cached || fetchPromise;
    })
  );
});

