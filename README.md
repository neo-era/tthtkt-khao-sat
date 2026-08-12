# Khảo sát chiếu sáng khu vực Trung Tâm

Ứng dụng web khảo sát hiện trường (chạy trên điện thoại) để rà soát, đề xuất **bổ sung hệ thống chiếu sáng công cộng** tại **610 trạm dừng xe buýt** do Trung tâm Quản lý Giao thông công cộng quản lý, địa bàn **Quận 1, 3, 5, 8, 10, 11, Phú Nhuận, Bình Thạnh**.

> Nguồn danh mục: `docs/DS TRẠM DỪNG XE BUÝT.xlsx` — Danh sách trạm dừng xe buýt do Trung tâm Quản lý Giao thông công cộng (Sở Xây dựng TP.HCM) quản lý.

---

## 1. Bộ công cụ gồm những gì

| Thành phần | Tập tin | Công dụng |
|---|---|---|
| **Web app khảo sát** | `index.html` | Chạy trên điện thoại: nạp sẵn 610 vị trí, lấy GPS, chụp ảnh, lưu offline, xuất CSV, đồng bộ. |
| **Backend Google Sheets** | `apps-script/Code.gs` | Nhận dữ liệu từ app, lưu vào Google Sheet (có chống trùng theo STT). |
| **Danh mục nguồn** | `docs/DS TRẠM DỪNG XE BUÝT.xlsx` | File Excel gốc (3.754 dòng); `tools/build_raw.py` lọc ra 610 vị trí trong phạm vi và nạp vào `index.html`. |
| **Biểu mẫu Excel** | `docs/Bieu_mau_khao_sat_...xlsx` | ⚠️ Phiếu Excel cũ, **vẫn theo danh mục 107 vị trí đã bỏ** — cần dựng lại nếu còn dùng. |
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
   - Lần đầu chạy sẽ xin cả quyền **Google Drive** (để lưu ảnh hiện trường) — cứ Allow.
   - Chạy xong, sheet "KhaoSat" sẽ có sẵn 22 cột tiêu đề định dạng đẹp.
5. Bấm **Deploy → New deployment** → biểu tượng bánh răng chọn **Web app**:
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
   - → **Deploy** → **copy đường link kết thúc bằng `/exec`**.
6. Dán link `/exec` vào hằng số **`DEFAULT_URL`** trong `index.html` (app không còn ô nhập URL), rồi commit lại.

> Mỗi khi sửa `Code.gs`, phải **Deploy lại**: *Manage deployments → Edit (bút chì) → Version: New version → Deploy*. URL giữ nguyên.

**Kiểm tra nhanh:** mở link `/exec` trên trình duyệt sẽ thấy JSON trạng thái. Thêm `?action=data` vào cuối để xem toàn bộ dữ liệu đã đồng bộ (đây cũng chính là đường app dùng để đọc ngược trạng thái về máy).

---

## 4. Cách sử dụng ngoài hiện trường

1. Mở app (nên **Thêm vào màn hình chính** để dùng như app: trình duyệt → menu → *Add to Home Screen*).
2. Màn hình danh sách: lọc theo địa bàn (chip), chọn **phường/xã** ở ô bên dưới, lọc trạng thái hoặc gõ tìm kiếm; xem tiến độ X/610 ở trên cùng. Danh sách nạp dần 200 vị trí mỗi lần (nút **Xem thêm**).
3. Chạm một vị trí → bấm **Định vị** để lấy tọa độ GPS, nhập 2 khoảng cách (trụ đèn ↔ trụ dừng, trụ dừng ↔ trụ đèn kế), chọn **Loại trụ dừng/nhà chờ**, **Vị trí trụ so với nhà chờ**, **Loại trụ đèn**, **Ưu tiên đề xuất**, chụp **tối đa 3 ảnh**, ghi chú thêm → **Lưu khảo sát điểm này**. Bộ trường bám theo bảng giấy *Tổng hợp kết quả tính toán lựa chọn bộ đèn LED*.
   - **Bấm Định vị trước khi chụp ảnh** — ảnh được đóng dấu tọa độ ngay lúc chụp, chưa có tọa độ thì trên ảnh ghi "chưa lấy được".
4. Dữ liệu được **lưu ngay trên điện thoại** (đóng app mở lại không mất). Có mạng thì ảnh tự lên **Google Drive** trước, rồi dòng dữ liệu kèm link ảnh lên **Google Sheets**. Mỗi ảnh cũng tự lưu một bản về bộ nhớ máy.
5. Thẻ trong danh sách hiện **Đã lên Sheets** (xanh) hay **Chưa đồng bộ** (vàng) để biết điểm nào đã tới bảng tính.
6. Khi mất mạng lúc lưu: **không phải làm gì cả** — app tự gửi lại khi có mạng trở lại, khi mở app, và mỗi lần mở bảng **Xuất dữ liệu**.

---

## 5. Câu hỏi thường gặp

**Đồng bộ nhiều lần có bị trùng dữ liệu không?**
Không. Backend dùng cơ chế *upsert theo STT*: điểm đã có sẽ được cập nhật đè, không tạo dòng mới.

**Mất mạng có khảo sát được không?**
Được. App lưu offline; khi có mạng đồng bộ sau.

**Báo "Bộ nhớ đầy"?**
Do ảnh chưa kịp lên Drive tích lại trên máy. Ảnh đã nén xuống ~1024px, nhưng khảo sát offline nhiều điểm liên tiếp (mỗi điểm tối đa 3 ảnh) vẫn có thể đầy. Hãy vào **Xuất / Đồng bộ** khi có mạng — ảnh lên Drive xong là app tự xóa bản nặng khỏi bộ nhớ máy.

**Ảnh lưu ở đâu?**
Ba nơi: (1) bộ nhớ máy — tự tải về ngay khi chụp, đây là bản sao lưu chắc chắn nhất; (2) **Google Drive**, thư mục *Anh khao sat chieu sang - Tram dung xe buyt* trong Drive của tài khoản chạy Apps Script, tên file `KS_<STT>_<số ảnh>.jpg`; (3) **link** tới file Drive ghi vào 3 cột *Ảnh 1/2/3* của Google Sheet. Chụp lại cùng một ô ảnh sẽ ghi đè file cũ trên Drive.

**Ảnh có thông tin gì trên đó?**
Ngay lúc chụp, app vẽ đè lên góc dưới ảnh: **địa chỉ điểm** (tuyến đường – số nhà, phường/xã, địa bàn), **tọa độ GPS** kèm sai số, và **ngày giờ chụp**. Không cần mạng vì địa chỉ lấy từ danh mục có sẵn trong app.

**Nhiều cán bộ cùng khảo sát?**
Mỗi người mở cùng một link, đồng bộ về **cùng một Google Sheet**. Backend có khóa (LockService) chống ghi đè khi đồng bộ đồng thời.

App còn **đọc ngược từ Sheet** nên máy này thấy được điểm người khác đã làm: mỗi lần mở app, khi có mạng trở lại, và khi mở bảng **Xuất dữ liệu**. Điểm nào máy anh đang sửa mà chưa gửi lên được thì **không bị đè** — bản trên máy luôn được ưu tiên cho tới khi gửi thành công.

**Muốn nhập liệu trên máy tính?**
Nhập thẳng vào Google Sheet — app không còn chức năng nhập ngược từ CSV. CSV chỉ để **xuất** ra xem/lưu trữ (đủ 610 dòng, có BOM nên Excel đọc đúng dấu). Biểu mẫu `docs/Bieu_mau_khao_sat_...xlsx` vẫn theo danh mục 107 vị trí cũ, chưa dựng lại theo danh mục mới.

**Dữ liệu khảo sát cũ còn không?**
Không dùng nữa. Danh mục đánh lại STT và bộ trường khảo sát cũng đổi theo bảng chọn bộ đèn LED, nên app chuyển sang khóa lưu `lavipco_ks_data_v3`. Bản ghi cũ vẫn nằm trong `localStorage` ở `lavipco_ks_data_v1` / `_v2` (không bị xóa) nếu cần tra lại.

**Đổi bộ trường thì Google Sheet cũ thế nào?**
Số cột đổi từ 20 thành 22 và tên cột khác hẳn. Phải chạy lại `setupSheet` (hoặc `clearData` rồi sửa hàng tiêu đề) và **Deploy version mới**, nếu không dữ liệu sẽ lệch cột.

**Bấm Lưu rồi mà Sheet chưa có dòng?**
App gửi ngay khi bấm Lưu, nhưng chỉ đánh dấu **Đã lên Sheets** sau khi Apps Script xác nhận ghi xong — nên thẻ còn nhãn vàng nghĩa là chưa tới nơi thật. App sẽ tự gửi lại khi có mạng. Nếu mãi không lên: kiểm tra đã **Deploy version mới** sau khi sửa `Code.gs` chưa, và URL `/exec` trong app có đúng không.

---

## 6. Cấu trúc thư mục

```
chieu-sang-trung-tam/
├── index.html              # Web app khảo sát (đặt ở gốc để GitHub Pages chạy)
├── huongdan.html           # Trang hướng dẫn sử dụng
├── baocao.html             # Trang báo cáo kết quả (in / lưu PDF)
├── manifest.json           # Cấu hình "thêm vào màn hình chính"
├── .nojekyll               # Tắt xử lý Jekyll của GitHub Pages
├── icons/
│   ├── icon-192.png
│   └── icon-512.png
├── apps-script/
│   └── Code.gs             # Backend Google Apps Script
├── tools/
│   └── build_raw.py        # Nạp danh mục từ Excel vào index.html
└── docs/
    ├── DS TRẠM DỪNG XE BUÝT.xlsx                       # Danh mục nguồn (3.754 dòng, lọc còn 610)
    └── Bieu_mau_khao_sat_chieu_sang_mang_xanh_v1.0.xlsx # Biểu mẫu cũ (danh mục 107 vị trí)
```

---

*Chiếu sáng khu vực Trung Tâm — Tài liệu nội bộ phục vụ công tác khảo sát hạ tầng chiếu sáng đô thị.*
