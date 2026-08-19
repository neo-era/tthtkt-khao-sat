/**
 * ============================================================================
 *  CHIẾU SÁNG KHU VỰC TRUNG TÂM – BACKEND GOOGLE APPS SCRIPT
 *  Web app khảo sát bổ sung chiếu sáng công cộng
 *  Trạm dừng xe buýt địa bàn Quận 1, 3, 5, 8, 10, 11, Phú Nhuận, Bình Thạnh
 * ----------------------------------------------------------------------------
 *  CÁCH CÀI ĐẶT (làm 1 lần):
 *   1. Tạo 1 Google Sheet mới (đặt tên tùy ý).
 *   2. Vào menu: Tiện ích mở rộng (Extensions) → Apps Script.
 *   3. Xóa code mẫu, dán toàn bộ nội dung file này vào, bấm Lưu.
 *   4. Trên thanh hàm, chọn "setupSheet" → Chạy (Run) để tạo tiêu đề & định dạng.
 *      (Lần đầu sẽ hỏi cấp quyền → chọn tài khoản → Advanced → Allow.
 *       Script cần thêm quyền Google Drive để lưu ảnh hiện trường.)
 *   5. Bấm "Deploy" (Triển khai) → New deployment → chọn loại "Web app":
 *         - Execute as:  Me (chính bạn)
 *         - Who has access:  Anyone (Bất kỳ ai)
 *      → Deploy → Copy đường dẫn kết thúc bằng "/exec".
 *   6. Dán URL đó vào ô "URL Google Apps Script" trong app khảo sát trên điện thoại.
 *
 *  LƯU Ý: Mỗi lần sửa code, phải Deploy lại (Manage deployments → Edit → New version)
 *  thì URL mới có hiệu lực; hoặc dùng "Test deployments" cho URL /dev khi thử.
 * ============================================================================
 */

// Tên sheet chứa dữ liệu khảo sát. Đổi nếu muốn.
var SHEET_NAME = 'KhaoSat';

// Thư mục Google Drive chứa ảnh hiện trường (tự tạo trong Drive của tài khoản chạy script).
var PHOTO_FOLDER = 'Anh khao sat chieu sang - Tram dung xe buyt';

// 22 cột app gửi lên + 1 cột thời gian server nhận = 23 cột.
var HEADERS = [
  'STT', 'Mã điểm dừng', 'Địa bàn', 'Tên vị trí', 'Tuyến đường', 'Số nhà/Vị trí',
  'Phường/Xã', 'Hạ tầng', 'Vĩ độ', 'Kinh độ',
  'KC trụ đèn → trụ dừng (m)', 'KC trụ dừng → trụ đèn kế (m)',
  'Loại trụ dừng, nhà chờ', 'Vị trí trụ so với nhà chờ', 'Loại trụ đèn',
  'Ưu tiên đề xuất', 'Ghi chú thêm',
  'Ảnh 1', 'Ảnh 2', 'Ảnh 3',
  'Người khảo sát', 'Thời gian lưu (máy)', 'Thời gian nhận (server)'
];
var NCOLS = HEADERS.length;        // 23
var APP_COLS = 22;                 // số cột app gửi
var TZ = 'Asia/Ho_Chi_Minh';

/**
 * Nhận dữ liệu từ app khảo sát (POST JSON là mảng các dòng).
 * Có upsert theo STT (cột 1): điểm đã có thì cập nhật, chưa có thì thêm mới.
 */
function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonOut_({ ok: false, error: 'Không có dữ liệu gửi lên.' });
  }

  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (parseErr) {
    return jsonOut_({ ok: false, error: 'Dữ liệu gửi lên không phải JSON hợp lệ.' });
  }

  // Ảnh hiện trường: {action:'photo', stt, idx, mime, data(base64)} → lưu vào Drive, trả link.
  // Không cần khóa vì mỗi ảnh có tên file riêng theo STT + vị trí ô ảnh.
  if (payload && !Array.isArray(payload) && payload.action === 'photo') {
    return savePhoto_(payload);
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // tránh ghi đè khi nhiều người đồng bộ cùng lúc
  } catch (lockErr) {
    return jsonOut_({ ok: false, error: 'Hệ thống đang bận, vui lòng đồng bộ lại.' });
  }

  try {
    var rows = Array.isArray(payload) ? payload : [payload];

    var sheet = getSheet_();
    var sttToRow = buildSttIndex_(sheet);   // { '12': 5, ... } STT -> số dòng trên sheet
    var now = nowStr_();
    var inserted = 0, updated = 0;

    rows.forEach(function (r) {
      if (!Array.isArray(r)) return;
      r = r.slice(0, APP_COLS);              // chỉ lấy 21 cột đầu
      while (r.length < APP_COLS) r.push('');
      r.push(now);                            // cột 22: thời gian server nhận

      var stt = String(r[0]).trim();
      if (stt && sttToRow[stt]) {
        // Cập nhật dòng đã tồn tại
        sheet.getRange(sttToRow[stt], 1, 1, NCOLS).setValues([r]);
        updated++;
      } else {
        // Thêm dòng mới
        sheet.appendRow(r);
        if (stt) sttToRow[stt] = sheet.getLastRow();
        inserted++;
      }
    });

    SpreadsheetApp.flush();
    return jsonOut_({
      ok: true,
      inserted: inserted,
      updated: updated,
      total: Math.max(0, sheet.getLastRow() - 1),
      receivedAt: now
    });

  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * GET: trả về trạng thái nhanh (mở URL /exec trên trình duyệt để kiểm tra).
 * Thêm ?action=data để lấy toàn bộ dữ liệu dạng JSON.
 */
function doGet(e) {
  var sheet = getSheet_();
  var action = (e && e.parameter && e.parameter.action) || 'status';

  // Trả head + values (mảng theo đúng thứ tự cột) để app đọc theo VỊ TRÍ cột,
  // không phụ thuộc chữ tiêu đề — đổi chữ tiêu đề không làm hỏng đồng bộ ngược.
  if (action === 'data') {
    var values = sheet.getDataRange().getValues();
    var head = values.shift() || [];
    return jsonOut_({
      ok: true,
      count: values.length,
      head: head,
      values: values.map(function (row) {
        return row.map(function (v) { return v === null || v === undefined ? '' : String(v); });
      })
    });
  }

  // Trả 1 ảnh dưới dạng base64 để trang báo cáo nhúng vào file Word.
  // Trình duyệt không đọc trực tiếp được file trên Drive (không có CORS),
  // nên phải đi vòng qua đây — Apps Script thì gọi được từ trang web.
  if (action === 'img') {
    var id = String((e.parameter && e.parameter.id) || '').trim();
    var w = Number((e.parameter && e.parameter.w) || 0);
    if (!id) return jsonOut_({ ok: false, error: 'Thiếu id ảnh.' });
    try {
      var blob = null;

      // Có ?w= thì lấy bản thu nhỏ cho file Word nhẹ bớt (ảnh gốc ~1024px trong
      // khi bản in chỉ rộng ~195px). UrlFetchApp gọi được link thumbnail vì chạy
      // trên máy chủ, không vướng CORS như trình duyệt. Hỏng thì rơi về ảnh gốc.
      if (w > 0) {
        try {
          var res = UrlFetchApp.fetch(
            'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w' + w,
            { muteHttpExceptions: true, followRedirects: true,
              headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
          if (res.getResponseCode() === 200) {
            var tb = res.getBlob();
            if (String(tb.getContentType()).indexOf('image/') === 0) blob = tb;
          }
        } catch (thumbErr) { /* dùng ảnh gốc */ }
      }

      if (!blob) blob = DriveApp.getFileById(id).getBlob();

      return jsonOut_({
        ok: true,
        id: id,
        mime: blob.getContentType() || 'image/jpeg',
        data: Utilities.base64Encode(blob.getBytes())
      });
    } catch (imgErr) {
      return jsonOut_({ ok: false, error: String(imgErr) });
    }
  }

  return jsonOut_({
    ok: true,
    service: 'Chiếu sáng khu vực Trung Tâm – Khảo sát chiếu sáng',
    sheet: SHEET_NAME,
    total: Math.max(0, sheet.getLastRow() - 1),
    time: nowStr_()
  });
}

/**
 * Chạy 1 lần để khởi tạo tiêu đề + định dạng bảng cho đẹp & dễ đọc.
 */
function setupSheet() {
  var sheet = getSheet_();
  sheet.clear();

  // Hàng tiêu đề
  sheet.getRange(1, 1, 1, NCOLS).setValues([HEADERS]);
  var hd = sheet.getRange(1, 1, 1, NCOLS);
  hd.setFontFamily('Times New Roman')
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#0F2A43')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 40);

  // Độ rộng cột (đơn vị pixel)
  var widths = [45, 90, 95, 200, 130, 130, 110, 120, 95, 95, 110, 110, 150, 140, 120, 90, 200,
                190, 190, 190, 120, 130, 130];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  // Định dạng vùng dữ liệu
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, NCOLS)
       .setFontFamily('Times New Roman')
       .setVerticalAlignment('middle')
       .setWrap(true);

  // Tô màu xen kẽ dòng (banding) cho dễ nhìn
  try {
    var existing = sheet.getBandings();
    existing.forEach(function (b) { b.remove(); });
    sheet.getRange(1, 1, Math.max(2, sheet.getMaxRows()), NCOLS)
         .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  } catch (bandErr) { /* bỏ qua nếu môi trường không hỗ trợ */ }

  SpreadsheetApp.getActive().toast('Đã khởi tạo bảng "' + SHEET_NAME + '" với ' + NCOLS + ' cột.', 'Chiếu sáng khu vực Trung Tâm', 5);
}

/**
 * NÂNG CẤP BẢNG ĐANG CÓ DỮ LIỆU: chèn cột "Mã điểm dừng" vào vị trí thứ 2.
 * Chạy MỘT LẦN duy nhất, sau khi dán bản Code.gs mới này vào.
 *
 * ⚠ ĐỪNG dùng setupSheet để nâng cấp — hàm đó gọi sheet.clear(), mất sạch dữ liệu
 * khảo sát đã đồng bộ. insertColumnBefore giữ nguyên dữ liệu cũ, chỉ đẩy các cột
 * phía sau sang phải đúng một nấc cho khớp bố cục mới.
 *
 * Chạy lại lần hai không sao: thấy tiêu đề đã đúng thì bỏ qua.
 */
function themCotMaDiemDung() {
  var sheet = getSheet_();
  if (String(sheet.getRange(1, 2).getValue()).trim() === HEADERS[1]) {
    SpreadsheetApp.getActive().toast('Bảng đã có cột "Mã điểm dừng", không cần làm gì.',
      'Chiếu sáng khu vực Trung Tâm', 5);
    return;
  }

  sheet.insertColumnBefore(2);
  sheet.getRange(1, 1, 1, NCOLS).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, NCOLS)
       .setFontFamily('Times New Roman')
       .setFontWeight('bold')
       .setFontColor('#FFFFFF')
       .setBackground('#0F2A43')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setWrap(true);
  sheet.setColumnWidth(2, 90);
  SpreadsheetApp.flush();

  SpreadsheetApp.getActive().toast(
    'Đã chèn cột "Mã điểm dừng". Bảng giờ có ' + NCOLS + ' cột, dữ liệu cũ giữ nguyên.',
    'Chiếu sáng khu vực Trung Tâm', 6);
}

/**
 * Xóa toàn bộ dữ liệu khảo sát (giữ lại tiêu đề). Dùng khi cần làm lại từ đầu.
 */
function clearData() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, NCOLS).clearContent();
  SpreadsheetApp.getActive().toast('Đã xóa dữ liệu, giữ lại tiêu đề.', 'Chiếu sáng khu vực Trung Tâm', 4);
}

/* ------------------------------ ẢNH HIỆN TRƯỜNG ------------------------------ */

/**
 * Lưu 1 ảnh vào thư mục Drive, đặt quyền "ai có link cũng xem được" để mở từ Sheet.
 * Tên file cố định theo STT + ô ảnh (KS_0001_1.jpg) nên chụp lại sẽ ghi đè, không sinh rác.
 */
function savePhoto_(p) {
  try {
    var stt = String(p.stt || '').replace(/\D/g, '');
    var idx = Number(p.idx || 0) + 1;
    if (!stt) return jsonOut_({ ok: false, error: 'Thiếu STT của điểm khảo sát.' });
    if (!p.data) return jsonOut_({ ok: false, error: 'Thiếu dữ liệu ảnh.' });

    var name = 'KS_' + ('0000' + stt).slice(-4) + '_' + idx + '.jpg';
    var folder = getPhotoFolder_();

    // Xóa bản cũ cùng tên (nếu chụp lại ô ảnh này)
    var old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);

    var blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || 'image/jpeg', name);
    var file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) { /* miền hạn chế chia sẻ ra ngoài — vẫn giữ file, link nội bộ dùng được */ }

    return jsonOut_({
      ok: true,
      id: file.getId(),
      url: 'https://drive.google.com/file/d/' + file.getId() + '/view',
      name: name
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function getPhotoFolder_() {
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

/* ------------------------------ HÀM PHỤ TRỢ ------------------------------ */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  // Đảm bảo luôn có hàng tiêu đề
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, NCOLS).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function buildSttIndex_(sheet) {
  var map = {};
  var last = sheet.getLastRow();
  if (last < 2) return map;
  var col = sheet.getRange(2, 1, last - 1, 1).getValues(); // cột STT từ dòng 2
  for (var i = 0; i < col.length; i++) {
    var v = String(col[i][0]).trim();
    if (v) map[v] = i + 2; // số dòng thực tế trên sheet
  }
  return map;
}

function nowStr_() {
  return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
