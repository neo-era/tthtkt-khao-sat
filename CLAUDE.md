# CLAUDE.md

File ngữ cảnh cho Claude (Claude Code / Cowork) khi làm việc với repo này.
Đọc kỹ trước khi sửa code.

## Dự án là gì

Web app khảo sát hiện trường (mobile-first) để rà soát, đề xuất **bổ sung chiếu sáng công cộng** tại **610 trạm dừng xe buýt** do Trung tâm Quản lý Giao thông công cộng quản lý, **địa bàn Quận 1, 3, 5, 8, 10, 11, Phú Nhuận, Bình Thạnh**. Thuộc **Chiếu sáng khu vực Trung Tâm**.

Nguồn danh mục: `docs/DS TRẠM DỪNG XE BUÝT.xlsx` (Trung tâm Quản lý Giao thông công cộng – Sở Xây dựng TP.HCM). Danh mục 107 vị trí theo CV 3471/TTGTHTKT-CXCS1 **đã bỏ** từ 12/8/2026.

Đối tượng dùng: cán bộ khảo sát đi thực địa bằng điện thoại (thường ngoài trời, có thể mất mạng).

## Kiến trúc (3 thành phần)

1. **Front-end** — `index.html`: một file HTML duy nhất, **vanilla JS thuần, KHÔNG có thư viện ngoài / CDN / build step**. Lưu offline bằng `localStorage`.
2. **Back-end** — `apps-script/Code.gs`: Google Apps Script Web App, lưu dữ liệu vào Google Sheet. Upsert theo STT.
3. **Danh mục nguồn** — `docs/DS TRẠM DỪNG XE BUÝT.xlsx` + `tools/build_raw.py` (cần `openpyxl`) để nạp lại mảng `RAW` vào `index.html`. File Excel có 3.754 dòng cho toàn TP.HCM/Vũng Tàu/Bình Dương; **phạm vi khảo sát chỉ là 610 dòng có mã địa bàn ở cột 6** (Q1/Q3/Q5/Q8/Q10/Q11/PN/BT), các dòng còn lại bị script bỏ qua.
4. **Biểu mẫu** — `docs/Bieu_mau_khao_sat_...xlsx`: phiên bản Excel để in / nhập trên máy tính. ⚠️ **Chưa cập nhật** — vẫn theo danh mục 107 vị trí cũ.

Luồng: app (điện thoại) → lưu localStorage → khi có mạng POST JSON lên Apps Script → ghi vào Google Sheet. Hoặc xuất/nhập CSV.

## RÀNG BUỘC BẮT BUỘC (đừng phá vỡ)

- **Không thêm dependency ngoài.** App phải chạy offline ngoài hiện trường → không thêm `<script src>` từ CDN, không npm, không framework. Mọi thứ inline trong `index.html`. (Trang dashboard bản đồ tương lai nếu dùng Leaflet thì để **trang riêng**, không nhét vào app khảo sát chính.)
- **GPS & camera chỉ chạy qua `https://`** (GitHub Pages). Mở file trực tiếp từ máy thì hai tính năng bị chặn — đây là hành vi của trình duyệt, không phải bug.
- **Upsert theo STT**: app đồng bộ lại TOÀN BỘ điểm "đã khảo sát" mỗi lần. Backend phải cập nhật đè theo cột STT, không được `appendRow` mù → sẽ nhân đôi dữ liệu. Logic này nằm ở `doPost` + `buildSttIndex_` trong `Code.gs`.
- **Tiếng Việt có dấu**: giữ UTF-8. CSV export phải có BOM `\ufeff` để Excel đọc đúng dấu.
- **Số cột phải khớp**: app gửi đúng **21 cột** theo `rowArray()`; backend thêm cột 22 (thời gian server). Nếu đổi schema phải sửa ĐỒNG THỜI `rowArray` + `HEADERS` trong `index.html` và `HEADERS`/`APP_COLS`/`NCOLS`/`widths` trong `Code.gs`.
- **Tiêu đề CSV phải qua `csvCell`**: `HEADERS` có ô chứa dấu phẩy ("Loại trụ dừng, nhà chờ"), `HEADERS.join(",")` trần sẽ làm hàng tiêu đề nhiều hơn hàng dữ liệu 1 cột và lệch toàn bộ khi mở bằng Excel.

## Schema dữ liệu

### Bản ghi khảo sát (object trong localStorage `lavipco_ks_data_v3`)
Bộ trường bám theo bảng giấy **"Tổng hợp kết quả tính toán lựa chọn bộ đèn LED"**.
Key = STT (số). Value:
```js
{ lat, lng, acc, gpsAt, d1, d2,
  nhacho, vitri, loaitru, prio, note, by, photos, at, sync }
```
- `d1` = khoảng cách trụ đèn → trụ dừng/nhà chờ (m, theo hướng xe buýt lưu thông); `d2` = khoảng cách trụ dừng/nhà chờ → trụ đèn tiếp theo (m). Cả hai nhập số tự do.
- `nhacho`  ∈ {Nhà chờ chưa cải tạo, Đã lắp mới}
- `vitri`   ∈ {Cùng phía, Đối diện}
- `loaitru` ∈ {BTLT, STK côn tròn, STK bát giác}  ← loại trụ đèn, không phải loại trạm dừng (đó là `RAW.loai`)
- `prio`    ∈ {Ưu tiên 1, Ưu tiên 2}
- `photos` = mảng **tối đa 3 ảnh**, mỗi phần tử `{ d, id, url }`:
  - `d` = dataURL JPEG đã đóng dấu (max 1024px, q0.6). **Chỉ giữ tới khi tải lên Drive xong** rồi xóa — nếu giữ cả 3 ảnh full cho hàng trăm điểm thì `localStorage` (~5MB) tràn ngay.
  - `id` / `url` = ID và link file trên Google Drive sau khi tải lên. Ảnh đã lên Drive hiển thị lại qua `https://drive.google.com/thumbnail?id=<id>&sz=w400` (cần mạng).
  - Ảnh được vẽ lại bằng canvas: thu nhỏ rồi in đè **địa chỉ · tọa độ (kèm ±độ chính xác) · thời gian chụp** ở dải dưới (`drawStamp` / `stampLines`). Tọa độ lấy từ ô Vĩ độ/Kinh độ đang nhập — chưa bấm Định vị thì đóng dấu "chưa lấy được".
  - Mỗi ảnh vẫn tự tải 1 bản về máy ngay khi chụp (đó là bản sao lưu thật, Drive chỉ là nơi chia sẻ).
- `sync` = thời điểm Apps Script **xác nhận đã ghi** dòng này vào Sheet. Rỗng = chưa lên Sheets → thẻ hiện nhãn vàng "Chưa đồng bộ", `unsynced()` gom lại để gửi tiếp. `saveDetail` xóa `sync` mỗi lần lưu vì nội dung đã đổi.
- Một điểm coi là "đã khảo sát" (`isDone`) khi có một trong: lat, d1, d2, nhacho, vitri, loaitru, prio, note, photos.

### Thứ tự 21 cột gửi lên (hàm `rowArray` trong index.html)
`STT, Địa bàn, Tên vị trí, Tuyến đường, Số nhà, Phường/Xã, Hạ tầng, Vĩ độ, Kinh độ, KC trụ đèn→trụ dừng(m), KC trụ dừng→trụ đèn kế(m), Loại trụ dừng nhà chờ, Vị trí trụ so với nhà chờ, Loại trụ đèn, Ưu tiên đề xuất, Ghi chú thêm, Ảnh 1, Ảnh 2, Ảnh 3, Người KS, Thời gian lưu`

Cột 22 (server): `Thời gian nhận (server)` do `Code.gs` tự thêm.
Cột 17–19 là **link Drive**, chỉ xuất ra — `importCsv` bỏ qua (không dựng lại được file ảnh từ link).

### Đồng bộ Google Sheets
- URL Apps Script được nhúng sẵn (`DEFAULT_URL` trong index.html); ô nhập URL pre-fill bằng URL này.
- Mỗi lần bấm **Lưu khảo sát điểm này** → ngoài lưu localStorage còn **tải ảnh chưa lên Drive trước**, rồi mới POST dòng dữ liệu (đã kèm link ảnh) lên Sheets, nếu `navigator.onLine`. Offline thì lưu máy, đồng bộ sau ở mục Xuất/Đồng bộ (nút đó cũng tải ảnh tồn trước rồi mới gửi dòng).
- **`postRows` phải đọc JSON trả về và ném lỗi khi `ok === false`.** Chỉ dựa vào `fetch` thành công là sai: Apps Script trả HTTP 200 kèm `{ok:false}` khi lỗi (khóa bận, sai schema…), app sẽ báo "đã đồng bộ" trong khi Sheet trống. Đánh dấu `sync` chỉ diễn ra sau khi đã qua kiểm tra này.
- **Đồng bộ hai chiều** (`syncBothWays`): gửi phần còn tồn (`pushPending`) rồi đọc về (`pullFromSheets`). Chạy khi có sự kiện `online`, một lần lúc mở app, và mỗi lần mở bảng Xuất dữ liệu. Không có hàng đợi riêng — `unsynced()` suy ra từ `sync` rỗng nên không bao giờ lệch trạng thái.

### Đọc ngược từ Sheets (`pullFromSheets`)
Nhiều cán bộ dùng chung một Sheet → máy này phải thấy điểm máy kia đã khảo sát. Đọc `?action=data`, trộn vào `store` theo **vị trí cột** (`values` + `head` do `doGet` trả về dạng mảng, không phải object) — đổi chữ tiêu đề trên Sheet cũng không làm hỏng.

Hai luật **không được phá**, nếu không sẽ mất dữ liệu hiện trường:
1. Bản ghi cục bộ `isDone` mà `sync` rỗng = có sửa đổi chưa lên Sheets → **bỏ qua, không đè**.
2. Bản ghi còn ảnh `d` chưa có `id` (chưa lên Drive) → **bỏ qua**, vì Sheet chỉ có link, đè vào là mất ảnh.

Ngoài ra: dòng rỗng trên Sheet bị bỏ qua (không tạo bản ghi rác), và bản ghi không đổi thì không ghi lại (so bằng `JSON.stringify`) để khỏi báo nhầm "đã cập nhật".
- **App không còn nút đồng bộ tay / nhập CSV / xóa toàn bộ / ô nhập URL** (bỏ theo yêu cầu 12/8/2026 để màn hình gọn cho cán bộ hiện trường). Hệ quả: `currentSyncUrl()` chỉ đọc `LS_URL` rồi `DEFAULT_URL`; muốn đổi backend phải sửa `DEFAULT_URL` trong code hoặc set `LS_URL` qua console. Đường thoát duy nhất khi Sheets hỏng là **Tải file CSV**.

### Ảnh trên Google Drive
- Cùng một Web App URL, phân biệt bằng nội dung POST: mảng → ghi Sheet; object `{action:'photo', stt, idx, mime, data}` → `savePhoto_` lưu Drive.
- Thư mục `PHOTO_FOLDER` trong `Code.gs`, tự tạo lần đầu. Tên file **cố định** `KS_<stt 4 số>_<ô ảnh>.jpg` → chụp lại thì ghi đè, không sinh rác; nhờ vậy tra bằng `getFilesByName` nên không phải quét cả thư mục.
- File được `setSharing(ANYONE_WITH_LINK, VIEW)` để link trong Sheet mở được. Nếu tổ chức chặn chia sẻ ra ngoài thì lệnh này nuốt lỗi, file vẫn lưu — chỉ người trong tổ chức xem được.
- Giữ `Content-Type: text/plain` khi POST để trình duyệt **không preflight** (Apps Script không trả lời OPTIONS). App có đọc JSON trả về để lấy `id`/`url`.

### Các key localStorage
- `lavipco_ks_data_v3` — toàn bộ bản ghi khảo sát. Khóa `_v1` (danh mục 107 vị trí) và `_v2` (bộ trường chiếu sáng cũ) còn nằm trong localStorage nhưng không dùng — cả STT lẫn tên trường đều không tương ứng.
- `lavipco_ks_url_v1`  — URL Apps Script đè `DEFAULT_URL` (không còn UI để đặt, chỉ set qua console).
- `lavipco_ks_by`      — tên người khảo sát gần nhất (điền sẵn).
- `lavipco_ks_raw_v1`  — bản sao danh mục `RAW` do `index.html` ghi mỗi lần mở; `baocao.html` đọc lại để dựng báo cáo mà không cần nhúng lại RAW.

### Trang báo cáo (`baocao.html`)
Trang in được độc lập, đọc `lavipco_ks_raw_v1` + `lavipco_ks_data_v3` từ localStorage (cùng origin), tự tính tổng quan / phân bố theo địa bàn–phường/xã / thống kê loại nhà chờ – vị trí trụ – loại trụ đèn / khoảng cách trung bình + mức ưu tiên / bảng chi tiết nhóm theo phường/xã / phụ lục vị trí chưa khảo sát / kết luận. Vào từ nút "Xuất báo cáo khảo sát" trong sheet Xuất/Đồng bộ. **Phải mở `index.html` trước ít nhất 1 lần** để có `lavipco_ks_raw_v1`.

### Danh mục vị trí gốc (nguồn dữ liệu chuẩn)
Mảng `RAW` (JSON, 1 dòng duy nhất) nhúng trong `<script>` của `index.html`. Mỗi phần tử:
`{ stt, diaban, ten, duong, sonha, phuong, loai, ghichu }`.
- `diaban` = địa bàn quận: `Quận 1` / `Quận 3` / `Quận 5` / `Quận 8` / `Quận 10` / `Quận 11` / `Phú Nhuận` / `Bình Thạnh` (dùng cho chip lọc). Suy ra từ mã ở cột 6 file Excel, vì danh mục mới chỉ còn tên phường sau sáp nhập.
- `phuong` = phường/xã đã bỏ tiền tố "Phường/Xã/Thị trấn" (dùng cho ô chọn + tiêu đề nhóm).
- `ten` = `duong – sonha` (nhãn dễ đọc, chỉ dùng cho CSV/báo cáo).
- `loai` = hạ tầng (Trụ dừng loại 1/2/3/4/4S/5/5S, Trụ dừng, Trụ biển báo).
- `ghichu` luôn rỗng (nguồn mới không có cột này) nhưng giữ lại để không phá schema.
- `stt` đánh lại 1…610 **sau khi** đã sắp theo địa bàn → phường/xã, nên STT tăng dần đúng theo thứ tự hiển thị. STT **không** khớp với dòng trong file Excel (cột STT trong Excel là công thức SUBTOTAL, giá trị cache đã hỏng).

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

- Trang `map.html` riêng: đọc `?action=data` từ Apps Script, vẽ 610 điểm lên Leaflet + OSM, lọc theo trạng thái/ưu tiên.
- Dựng lại biểu mẫu Excel theo danh mục mới (610 vị trí).
- Nút xuất Excel (.xlsx) trực tiếp từ app thay vì CSV.
- Báo cáo tổng hợp tự động theo địa bàn.
