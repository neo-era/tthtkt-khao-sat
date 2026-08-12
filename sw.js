/* ============================================================================
 *  Service worker — cho phép cài app (PWA) trên điện thoại & máy tính,
 *  và mở được app khi ngoài hiện trường không có sóng.
 *
 *  Chiến lược: MẠNG TRƯỚC, CACHE DỰ PHÒNG (network-first).
 *  Lý do: app là 1 file HTML nhúng sẵn danh mục 610 vị trí, sửa đổi được đẩy
 *  lên GitHub Pages liên tục. Nếu ưu tiên cache thì cán bộ sẽ kẹt ở bản cũ mà
 *  không biết. Mỗi lần mở có sóng là tự lấy bản mới và cập nhật cache.
 *
 *  ĐỔI CODE THÌ NHỚ TĂNG SỐ Ở "VERSION" để cache cũ được dọn.
 * ==========================================================================*/
var VERSION = 'ks-v1';

var SHELL = [
  './',
  './index.html',
  './baocao.html',
  './huongdan.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(VERSION)
      // addAll thất bại toàn bộ nếu 1 file lỗi → nạp từng file, thiếu file nào bỏ qua file đó
      .then(function (c) { return Promise.all(SHELL.map(function (u) { return c.add(u).catch(function () {}); })); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys()
      .then(function (ks) {
        return Promise.all(ks.map(function (k) { return k === VERSION ? null : caches.delete(k); }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var req = e.request;

  // Chỉ đụng vào GET của chính trang này. POST lên Apps Script, ảnh trên Drive,
  // lệnh đọc ?action=data… phải đi thẳng ra mạng, tuyệt đối không cache.
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  e.respondWith(
    fetch(req)
      .then(function (res) {
        if (res && res.ok && res.type === 'basic') {
          var copy = res.clone();
          caches.open(VERSION).then(function (c) { c.put(req, copy); });
        }
        return res;
      })
      .catch(function () {
        return caches.match(req).then(function (hit) {
          if (hit) return hit;
          // Mở thẳng một đường dẫn lạ khi offline → trả về app chính
          if (req.mode === 'navigate') return caches.match('./index.html');
          return new Response('', { status: 504, statusText: 'Offline' });
        });
      })
  );
});
