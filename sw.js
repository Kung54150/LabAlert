// Service worker ของ LabAlert: ทำให้ติดตั้งเป็นแอปได้ และโหลดหน้าล่าสุดจากเน็ตก่อนเสมอ (ไม่ค้างเวอร์ชันเก่า)
// ไม่แตะคำขอไป Apps Script / LINE / CDN (ข้ามโดเมน) เพื่อไม่ให้ข้อมูลผู้ป่วยถูกเก็บในแคช
const CACHE = 'labalert-shell-v3';
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
    fetch(req, { cache: 'no-cache' }).then((res) => {   // ตรวจกับเซิร์ฟเวอร์ทุกครั้ง (ไม่ใช้ของเก่าในแคชเบราว์เซอร์ที่ GitHub Pages ให้เก็บได้ 10 นาที)
      if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req).then((hit) => hit || caches.match('index.html')))
  );
});

// กดที่การแจ้งเตือน: ดึงหน้าแอปที่เปิดอยู่ขึ้นมา (หรือเปิดใหม่) แล้วไปที่รายการนั้น
self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  const hn = (e.notification.data && e.notification.data.hn) || '';
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((list) => {
    for (const c of list) { if ('focus' in c) { c.postMessage({ type: 'open-case', hn: hn }); return c.focus(); } }
    return self.clients.openWindow('./?mode=follow&hn=' + encodeURIComponent(hn));
  }));
});
