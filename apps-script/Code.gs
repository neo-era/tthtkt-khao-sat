/**
 * ============================================================================
 *  CHIẾU SÁNG KHU VỰC TRUNG TÂM – BACKEND GOOGLE APPS SCRIPT
 *  Web app khảo sát bổ sung chiếu sáng công cộng (Phụ lục 107 vị trí)
 *  CV số 3471/TTGTHTKT-CXCS1 ngày 22/6/2026
 * ----------------------------------------------------------------------------
 *  CÁCH CÀI ĐẶT (làm 1 lần):
 *   1. Tạo 1 Google Sheet mới (đặt tên tùy ý).
 *   2. Vào menu: Tiện ích mở rộng (Extensions) → Apps Script.
 *   3. Xóa code mẫu, dán toàn bộ nội dung file này vào, bấm Lưu.
 *   4. Trên thanh hàm, chọn "setupSheet" → Chạy (Run) để tạo tiêu đề & định dạng.
 *      (Lần đầu sẽ hỏi cấp quyền → chọn tài khoản → Advanced → Allow.)
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

// 19 cột app gửi lên + 1 cột thời gian server nhận = 20 cột.
var HEADERS = [
  'STT', 'Mã điểm dừng', 'Tên điểm dừng', 'Tuyến đường', 'Số nhà/Vị trí',
  'Phường/Xã', 'Loại kết cấu', 'Vĩ độ', 'Kinh độ',
  'HT chiếu sáng', 'Nguồn điện (m)',
  'ĐX đèn', 'ĐX cáp', 'ĐX trụ', 'ĐX phụ kiện', 'Ưu tiên', 'Ghi chú khảo sát',
  'Người khảo sát', 'Thời gian lưu (máy)', 'Thời gian nhận (server)'
];
var NCOLS = HEADERS.length;        // 20
var APP_COLS = 19;                 // số cột app gửi
var TZ = 'Asia/Ho_Chi_Minh';

/**
 * Nhận dữ liệu từ app khảo sát (POST JSON là mảng các dòng).
 * Có upsert theo STT (cột 1): điểm đã có thì cập nhật, chưa có thì thêm mới.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // tránh ghi đè khi nhiều người đồng bộ cùng lúc
  } catch (lockErr) {
    return jsonOut_({ ok: false, error: 'Hệ thống đang bận, vui lòng đồng bộ lại.' });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut_({ ok: false, error: 'Không có dữ liệu gửi lên.' });
    }

    var rows = JSON.parse(e.postData.contents);
    if (!Array.isArray(rows)) rows = [rows];

    var sheet = getSheet_();
    var sttToRow = buildSttIndex_(sheet);   // { '12': 5, ... } STT -> số dòng trên sheet
    var now = nowStr_();
    var inserted = 0, updated = 0;

    rows.forEach(function (r) {
      if (!Array.isArray(r)) return;
      r = r.slice(0, APP_COLS);              // chỉ lấy 18 cột đầu
      while (r.length < APP_COLS) r.push('');
      r.push(now);                            // cột 19: thời gian server nhận

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

  if (action === 'data') {
    var values = sheet.getDataRange().getValues();
    var head = values.shift() || [];
    var out = values.map(function (row) {
      var o = {};
      head.forEach(function (h, i) { o[h] = row[i]; });
      return o;
    });
    return jsonOut_({ ok: true, count: out.length, rows: out });
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
  var widths = [45, 90, 200, 130, 130, 110, 110, 95, 95, 95, 80, 140, 140, 140, 140, 80, 200, 120, 130, 130];
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
 * Xóa toàn bộ dữ liệu khảo sát (giữ lại tiêu đề). Dùng khi cần làm lại từ đầu.
 */
function clearData() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, NCOLS).clearContent();
  SpreadsheetApp.getActive().toast('Đã xóa dữ liệu, giữ lại tiêu đề.', 'Chiếu sáng khu vực Trung Tâm', 4);
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
