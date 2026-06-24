# Khảo sát chiếu sáng khu vực Trung Tâm

Ứng dụng web khảo sát hiện trường (chạy trên điện thoại) để rà soát, đề xuất **bổ sung hệ thống chiếu sáng công cộng** tại **107 vị trí** trụ dừng / nhà chờ xe buýt và khu vực người đi bộ thuộc khu vực trung tâm TP. Hồ Chí Minh.

> Căn cứ: Công văn số 3471/TTGTHTKT-CXCS1 ngày 22/6/2026 và Phụ lục danh mục 107 vị trí của Trung tâm Quản lý Giao thông công cộng (Sở Xây dựng TP.HCM).

---

## 1. Bộ công cụ gồm những gì

| Thành phần | Tập tin | Công dụng |
|---|---|---|
| **Web app khảo sát** | `index.html` | Chạy trên điện thoại: nạp sẵn 107 vị trí, lấy GPS, chụp ảnh, lưu offline, xuất CSV, đồng bộ. |
| **Backend Google Sheets** | `apps-script/Code.gs` | Nhận dữ liệu từ app, lưu vào Google Sheet (có chống trùng theo STT). |
| **Biểu mẫu Excel** | `docs/Bieu_mau_khao_sat_...xlsx` | Phiếu khảo sát giấy/Excel để in hoặc nhập liệu trên máy tính. |
| Manifest + icon | `manifest.json`, `icons/` | Cho phép "Thêm vào màn hình chính" như một app thật. |

Toàn bộ app là **một file HTML duy nhất, không phụ thuộc thư viện ngoài**, nên chạy được cả khi mất mạng ngoài hiện trường.

---

## 2. Đăng web app lên GitHub Pages (≈ 5 phút)

### Cách A — Dùng giao diện web GitHub (không cần cài Git)

1. Đăng nhập [github.com](https://github.com) → bấm **New repository**.
2. Đặt tên, ví dụ `chieu-sang-trung-tam` → chọn **Public** → **Create repository**.
3. Trong repo trống, bấm **uploading an existing file** → kéo thả **toàn bộ nội dung bên trong thư mục này** (gồm `index.html`, `manifest.json`, thư mục `icons/`, `apps-script/`, `docs/`, file `.nojekyll`).
   - Lưu ý: kéo *các file bên trong*, không kéo nguyên thư mục mẹ, để `index.html` nằm ở **gốc repo**.
4. Bấm **Commit changes**.
5. Vào **Settings → Pages**:
   - **Source**: `Deploy from a branch`
   - **Branch**: `main` — thư mục `/ (root)` → **Save**.
6. Chờ 1–2 phút, tải lại trang Settings → Pages sẽ hiện đường dẫn:
   `https://<tên-tài-khoản>.github.io/chieu-sang-trung-tam/`
7. Mở đường dẫn đó trên điện thoại là dùng được.

### Cách B — Dùng Git (dòng lệnh)

```bash
cd chieu-sang-trung-tam
git init
git add .
git commit -m "Khởi tạo app khảo sát chiếu sáng khu vực Trung Tâm"
git branch -M main
git remote add origin https://github.com/<tên-tài-khoản>/chieu-sang-trung-tam.git
git push -u origin main
```

Rồi vào **Settings → Pages** bật như bước 5 ở trên.

> **Quan trọng:** GPS và camera **chỉ hoạt động khi mở qua `https://`** (tức là qua GitHub Pages). Nếu mở trực tiếp file `index.html` từ bộ nhớ máy thì hai tính năng này sẽ bị chặn — chỉ nhập tay được.

---

## 3. Dựng backend Google Sheets (làm 1 lần)

1. Tạo một **Google Sheet** mới (đặt tên tùy ý, ví dụ "Khảo sát chiếu sáng khu vực Trung Tâm 2026").
2. Menu **Tiện ích mở rộng (Extensions) → Apps Script**.
3. Xóa code mẫu, mở file `apps-script/Code.gs`, **copy toàn bộ** dán vào, bấm **Lưu**.
4. Trên thanh chọn hàm, chọn **`setupSheet`** → bấm **Run**.
   - Lần đầu sẽ hỏi cấp quyền: chọn tài khoản → **Advanced** → **Go to … (unsafe)** → **Allow**.
   - Chạy xong, sheet "KhaoSat" sẽ có sẵn 20 cột tiêu đề định dạng đẹp.
5. Bấm **Deploy → New deployment** → biểu tượng bánh răng chọn **Web app**:
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
   - → **Deploy** → **copy đường link kết thúc bằng `/exec`**.
6. Mở app khảo sát trên điện thoại → bấm **Xuất / Đồng bộ** → dán link `/exec` vào ô **URL Google Apps Script** → bấm **Đồng bộ lên Google Sheets**.

> Mỗi khi sửa `Code.gs`, phải **Deploy lại**: *Manage deployments → Edit (bút chì) → Version: New version → Deploy*. URL giữ nguyên.

**Kiểm tra nhanh:** mở link `/exec` trên trình duyệt sẽ thấy JSON trạng thái. Thêm `?action=data` vào cuối để xem toàn bộ dữ liệu đã đồng bộ.

---

## 4. Cách sử dụng ngoài hiện trường

1. Mở app (nên **Thêm vào màn hình chính** để dùng như app: trình duyệt → menu → *Add to Home Screen*).
2. Màn hình danh sách: tìm/lọc theo quận hoặc trạng thái, xem tiến độ X/107 ở trên cùng.
3. Chạm một vị trí → bấm **Định vị** để lấy tọa độ GPS, chọn hiện trạng chiếu sáng, nhập đề xuất bổ sung (Đèn / Cáp / Trụ / Phụ kiện), mức ưu tiên, chụp ảnh, ghi chú → **Lưu khảo sát điểm này**.
4. Dữ liệu được **lưu ngay trên điện thoại** (đóng app mở lại không mất) và **tự đồng bộ lên Google Sheets** nếu có mạng. Ảnh chụp được tự lưu một bản về bộ nhớ máy.
5. Khi mất mạng lúc lưu: vào **Xuất / Đồng bộ** → **Đồng bộ lên Google Sheets** khi có mạng, hoặc **Tải CSV** để mở bằng Excel.

---

## 5. Câu hỏi thường gặp

**Đồng bộ nhiều lần có bị trùng dữ liệu không?**
Không. Backend dùng cơ chế *upsert theo STT*: điểm đã có sẽ được cập nhật đè, không tạo dòng mới.

**Mất mạng có khảo sát được không?**
Được. App lưu offline; khi có mạng đồng bộ sau.

**Báo "Bộ nhớ đầy"?**
Do ảnh chụp tích nhiều. Hãy **Đồng bộ** hoặc **Tải CSV** rồi có thể xóa bớt; app đã tự nén ảnh xuống ~720px để tiết kiệm.

**Nhiều cán bộ cùng khảo sát?**
Mỗi người mở cùng một link, đồng bộ về **cùng một Google Sheet**. Backend có khóa (LockService) chống ghi đè khi đồng bộ đồng thời.

**Muốn nhập liệu trên máy tính?**
Dùng `docs/Bieu_mau_khao_sat_...xlsx` (đã nạp sẵn 107 vị trí, có ô chọn nhanh và bảng tổng hợp tự động).

---

## 6. Cấu trúc thư mục

```
chieu-sang-trung-tam/
├── index.html              # Web app khảo sát (đặt ở gốc để GitHub Pages chạy)
├── manifest.json           # Cấu hình "thêm vào màn hình chính"
├── .nojekyll               # Tắt xử lý Jekyll của GitHub Pages
├── icons/
│   ├── icon-192.png
│   └── icon-512.png
├── apps-script/
│   └── Code.gs             # Backend Google Apps Script
└── docs/
    └── Bieu_mau_khao_sat_chieu_sang_mang_xanh_v1.0.xlsx
```

---

*Chiếu sáng khu vực Trung Tâm — Tài liệu nội bộ phục vụ công tác khảo sát hạ tầng chiếu sáng đô thị.*
