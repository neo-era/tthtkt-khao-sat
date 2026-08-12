# CLAUDE.md

File ngữ cảnh cho Claude (Claude Code / Cowork) khi làm việc với repo này.
Đọc kỹ trước khi sửa code.

## Dự án là gì

Web app khảo sát hiện trường (mobile-first) để rà soát, đề xuất **bổ sung chiếu sáng công cộng** tại **3.754 trạm dừng xe buýt** do Trung tâm Quản lý Giao thông công cộng quản lý (TP.HCM 3.146 · Vũng Tàu 347 · Bình Dương 261). Thuộc **Chiếu sáng khu vực Trung Tâm**.

Nguồn danh mục: `docs/DS TRẠM DỪNG XE BUÝT.xlsx` (Trung tâm Quản lý Giao thông công cộng – Sở Xây dựng TP.HCM). Danh mục 107 vị trí theo CV 3471/TTGTHTKT-CXCS1 **đã bỏ** từ 12/8/2026.

Đối tượng dùng: cán bộ khảo sát đi thực địa bằng điện thoại (thường ngoài trời, có thể mất mạng).

## Kiến trúc (3 thành phần)

1. **Front-end** — `index.html`: một file HTML duy nhất, **vanilla JS thuần, KHÔNG có thư viện ngoài / CDN / build step**. Lưu offline bằng `localStorage`.
2. **Back-end** — `apps-script/Code.gs`: Google Apps Script Web App, lưu dữ liệu vào Google Sheet. Upsert theo STT.
3. **Danh mục nguồn** — `docs/DS TRẠM DỪNG XE BUÝT.xlsx` + `tools/build_raw.py` (cần `openpyxl`) để nạp lại mảng `RAW` vào `index.html`.
4. **Biểu mẫu** — `docs/Bieu_mau_khao_sat_...xlsx`: phiên bản Excel để in / nhập trên máy tính. ⚠️ **Chưa cập nhật** — vẫn theo danh mục 107 vị trí cũ.

Luồng: app (điện thoại) → lưu localStorage → khi có mạng POST JSON lên Apps Script → ghi vào Google Sheet. Hoặc xuất/nhập CSV.

## RÀNG BUỘC BẮT BUỘC (đừng phá vỡ)

- **Không thêm dependency ngoài.** App phải chạy offline ngoài hiện trường → không thêm `<script src>` từ CDN, không npm, không framework. Mọi thứ inline trong `index.html`. (Trang dashboard bản đồ tương lai nếu dùng Leaflet thì để **trang riêng**, không nhét vào app khảo sát chính.)
- **GPS & camera chỉ chạy qua `https://`** (GitHub Pages). Mở file trực tiếp từ máy thì hai tính năng bị chặn — đây là hành vi của trình duyệt, không phải bug.
- **Upsert theo STT**: app đồng bộ lại TOÀN BỘ điểm "đã khảo sát" mỗi lần. Backend phải cập nhật đè theo cột STT, không được `appendRow` mù → sẽ nhân đôi dữ liệu. Logic này nằm ở `doPost` + `buildSttIndex_` trong `Code.gs`.
- **Tiếng Việt có dấu**: giữ UTF-8. CSV export phải có BOM `\ufeff` để Excel đọc đúng dấu.
- **Số cột phải khớp**: app gửi đúng **19 cột** theo `rowArray()`; backend thêm cột 20 (thời gian server). Nếu đổi schema phải sửa ĐỒNG THỜI cả `rowArray`/`HEADERS` trong `index.html`, `HEADERS`/`APP_COLS`/`NCOLS`/`widths` trong `Code.gs`, đoạn mẫu `doPost` trong phần trợ giúp, và `importCsv` (chỉ số cột).

## Schema dữ liệu

### Bản ghi khảo sát (object trong localStorage `lavipco_ks_data_v2`)
Key = STT (số). Value:
```js
{ lat, lng, acc, gpsAt, light, power,
  den, cap, tru, phukien, prio, note, by, photo, at }
```
- `light` ∈ {Đầy đủ, Thiếu/Yếu, Hư hỏng, Không có}
- `den`, `cap`, `tru`, `phukien` = đề xuất bổ sung chiếu sáng tách 4 mục: đèn / cáp / trụ / phụ kiện (text tự do).
- `prio`  ∈ {Cao, Trung bình, Thấp}
- `photo` = dataURL JPEG đã nén (max 720px, q0.55). Có thể rỗng. Khi chụp/chọn ảnh được tự tải 1 bản về máy.
- Một điểm coi là "đã khảo sát" (`isDone`) khi có một trong: light, lat, note, den, cap, tru, phukien.

### Thứ tự 19 cột gửi lên (hàm `rowArray` trong index.html)
`STT, Khu vực, Tên vị trí, Tuyến đường, Số nhà, Phường/Xã, Hạ tầng, Vĩ độ, Kinh độ, HT chiếu sáng, Nguồn điện(m), ĐX đèn, ĐX cáp, ĐX trụ, ĐX phụ kiện, Ưu tiên, Ghi chú, Người KS, Thời gian lưu`

Cột 20 (server): `Thời gian nhận (server)` do `Code.gs` tự thêm.

### Đồng bộ Google Sheets
- URL Apps Script được nhúng sẵn (`DEFAULT_URL` trong index.html); ô nhập URL pre-fill bằng URL này.
- Mỗi lần bấm **Lưu khảo sát điểm này** → ngoài lưu localStorage còn POST riêng điểm đó lên Sheets (nếu `navigator.onLine`). Offline thì lưu máy, đồng bộ sau ở mục Xuất/Đồng bộ.

### Các key localStorage
- `lavipco_ks_data_v2` — toàn bộ bản ghi khảo sát (khóa `_v1` là dữ liệu của danh mục 107 vị trí cũ, giữ lại nhưng không dùng — STT hai danh mục không tương ứng nhau).
- `lavipco_ks_url_v1`  — URL Apps Script đã lưu.
- `lavipco_ks_by`      — tên người khảo sát gần nhất (điền sẵn).
- `lavipco_ks_raw_v1`  — bản sao danh mục `RAW` do `index.html` ghi mỗi lần mở; `baocao.html` đọc lại để dựng báo cáo mà không cần nhúng lại RAW.

### Trang báo cáo (`baocao.html`)
Trang in được độc lập, đọc `lavipco_ks_raw_v1` + `lavipco_ks_data_v2` từ localStorage (cùng origin), tự tính tổng quan / phân bố theo khu vực–phường/xã / thống kê hiện trạng / tổng hợp đề xuất theo đèn-cáp-trụ-phụ kiện / bảng chi tiết nhóm theo phường/xã / phụ lục vị trí chưa khảo sát (giới hạn 500 dòng đầu) / kết luận. Vào từ nút "Xuất báo cáo khảo sát" trong sheet Xuất/Đồng bộ. **Phải mở `index.html` trước ít nhất 1 lần** để có `lavipco_ks_raw_v1`.

### Danh mục vị trí gốc (nguồn dữ liệu chuẩn)
Mảng `RAW` (JSON, 1 dòng duy nhất) nhúng trong `<script>` của `index.html`. Mỗi phần tử:
`{ stt, diaban, ten, duong, sonha, phuong, loai, ghichu }`.
- `diaban` = khu vực: `TP.HCM` / `Vũng Tàu` / `Bình Dương` (dùng cho chip lọc).
- `phuong` = phường/xã đã bỏ tiền tố "Phường/Xã/Thị trấn" (dùng cho ô chọn + tiêu đề nhóm).
- `ten` = `duong – sonha` (nhãn dễ đọc, chỉ dùng cho CSV/báo cáo).
- `loai` = hạ tầng (Trụ dừng loại 1/2/3/4/4S/5/5S, Trụ dừng, Trụ biển báo).
- `ghichu` luôn rỗng (nguồn mới không có cột này) nhưng giữ lại để không phá schema.
- `stt` đánh lại 1…3754 **sau khi** đã sắp theo khu vực → phường/xã, nên STT tăng dần đúng theo thứ tự hiển thị. STT **không** khớp với dòng trong file Excel (cột STT trong Excel là công thức SUBTOTAL, giá trị cache đã hỏng).

Đây là **dữ liệu chỉ đọc**, không sửa khi khảo sát.

## Cách regenerate (nếu cần dựng lại từ nguồn)

```bash
pip install openpyxl
python tools/build_raw.py     # đọc docs/DS TRẠM DỪNG XE BUÝT.xlsx → thay dòng `const RAW = [...]` trong index.html
```

**Đừng sửa tay JSON trong `index.html`** — dễ sai dấu/escape. Nếu đổi danh mục: thay file Excel trong `docs/` rồi chạy lại script.

Biểu mẫu Excel (`docs/Bieu_mau_khao_sat_...xlsx`) trước đây do `build_form.py` sinh (script nằm ngoài repo) và **chưa được dựng lại** theo danh mục mới.

## Quy ước của Chiếu sáng khu vực Trung Tâm (áp dụng cho mọi file xuất)

- Đặt tên file: tăng số phiên bản mỗi lần (v1.0, v1.1…); ghi rõ đuôi định dạng; dùng **đuôi kép** (vd `...v1.1.xlsx.xlsx`) để đuôi còn hiển thị sau khi tải.
- Tài liệu pháp lý/chính thức: font **Times New Roman**. Biểu mẫu Excel đã dùng Times New Roman.
- Thông tin đơn vị (khi cần in lên tài liệu): **Chiếu sáng khu vực Trung Tâm**.

## Triển khai

- **GitHub Pages**: `index.html` ở GỐC repo; bật Settings → Pages → branch `main` / `/(root)`. File `.nojekyll` để tránh Jekyll bỏ qua file.
- **Apps Script**: chạy `setupSheet` 1 lần; Deploy Web app (Execute as Me, Access Anyone); sửa code phải Deploy version mới.

Chi tiết đầy đủ trong `README.md`.

## Việc có thể làm tiếp (gợi ý, chưa làm)

- Trang `map.html` riêng: đọc `?action=data` từ Apps Script, vẽ các điểm lên Leaflet + OSM (cần clustering ở quy mô 3.754 điểm), lọc theo trạng thái/ưu tiên.
- Dựng lại biểu mẫu Excel theo danh mục mới (tách file theo phường/xã, vì 3.754 dòng in ra không dùng được).
- Nút xuất Excel (.xlsx) trực tiếp từ app thay vì CSV.
- Báo cáo tổng hợp tự động theo địa bàn.
