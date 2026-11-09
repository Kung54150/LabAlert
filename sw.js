// Service worker ของ LabAlert: ทำให้ติดตั้งเป็นแอปได้ และโหลดหน้าล่าสุดจากเน็ตก่อนเสมอ (ไม่ค้างเวอร์ชันเก่า)
// ไม่แตะคำขอไป Apps Script / LINE / CDN (ข้ามโดเมน) เพื่อไม่ให้ข้อมูลผู้ป่วยถูกเก็บในแคช
const CACHE = 'labalert-shell-v1';
const SHELL = ['./', 'index.html', 'config.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).catch(() => {}).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  e.respondWith(
    fetch(req).then((res) => {
      if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req).then((hit) => hit || caches.match('index.html')))
  );
});
