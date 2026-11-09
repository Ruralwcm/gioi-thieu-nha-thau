/**
 * Google Apps Script — Backend for "Giới Thiệu Nhà Thầu"
 *
 * Deployment:
 * 1. Mở https://script.google.com/
 * 2. Tạo project mới
 * 3. Paste toàn bộ file này vào Code.gs
 * 4. Save (Ctrl+S)
 * 5. Deploy → New deployment → Web App
 *    - Execute as: Me
 *    - Access: Anyone
 * 6. Copy Web App URL → paste vào app phần Cấu hình GAS
 *
 * Data Sheet: 1RtVPThy95QMdHGvfL_9lCMNwcgStXCxxqLj1Cpf-KwY
 * Tab name: NhaThau_Data
 */

const SHEET_ID = '1RtVPThy95QMdHGvfL_9lCMNwcgStXCxxqLj1Cpf-KwY';
const TAB_NAME = 'NhaThau_Data';
const HEADERS = [
  'id', 'hangMuc', 'tenCty', 'nguoiDaiDien', 'sdt', 'tinhThanh', 'phuongXa', 'created_at'
];

/**
 * doGet — GET requests: ?action=list
 */
function doGet(e) {
  const params = e?.parameter || {};
  const action = params.action || 'list';

  if (action === 'list') {
    const data = getRecords();
    return jsonResponse({ success: true, data: data, count: data.length });
  }

  return jsonResponse({ success: false, error: 'Unknown action: ' + action });
}

/**
 * doPost — POST requests: { action: 'save', data: { ...record } }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData?.contents || '{}');
    const action = body.action;

    if (action === 'save') {
      const record = body.data;
      if (!record || !record.hangMuc || !record.tenCty) {
        return jsonResponse({ success: false, error: 'Thiếu thông tin bắt buộc (hạng mục, tên CTY)' });
      }
      saveRecord(record);
      return jsonResponse({ success: true, message: 'Đã lưu thông tin Nhà Thầu' });
    }

    return jsonResponse({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ success: false, error: 'Server error: ' + err.toString() });
  }
}

/**
 * Get or create the data sheet
 */
function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#C00000')
      .setFontColor('#FFFFFF');
    return sheet;
  }
  return sheet;
}

/**
 * Save a single record
 */
function saveRecord(record) {
  const sheet = getSheet();
  const id = Utilities.getUuid();
  const row = [
    id,
    record.hangMuc || '',
    record.tenCty || '',
    record.nguoiDaiDien || '',
    record.sdt || '',
    record.tinhThanh || '',
    record.phuongXa || '',
    record.created_at || new Date().toISOString()
  ];
  sheet.appendRow(row);
}

/**
 * Get all records
 */
function getRecords() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
  const values = range.getValues();

  return values.map(row => ({
    id:           String(row[0] || ''),
    hangMuc:      String(row[1] || ''),
    tenCty:       String(row[2] || ''),
    nguoiDaiDien: String(row[3] || ''),
    sdt:          String(row[4] || ''),
    tinhThanh:    String(row[5] || ''),
    phuongXa:     String(row[6] || ''),
    created_at:   String(row[7] || '')
  }));
}

/**
 * Helper: return JSON with CORS
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * (Optional) Initialize sheet — run once from Apps Script editor
 */
function initSheet() {
  const sheet = getSheet();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).clearContent();
  }
  Logger.log('Sheet initialized: ' + TAB_NAME);
}

/**
 * (Optional) Test: insert a sample record
 */
function testSave() {
  saveRecord({
    hangMuc: 'Cơ khí',
    tenCty: 'Công ty TNHH ABC',
    nguoiDaiDien: 'Nguyễn Văn A',
    sdt: '0912345678',
    tinhThanh: 'TP. Hồ Chí Minh',
    phuongXa: 'Phường Bến Nghé, Quận 1',
    created_at: new Date().toISOString()
  });
  Logger.log('Test record saved!');
}