/**
 * ETF 存股追蹤 — Google Apps Script
 *
 * 使用方式：
 * 1. 開一個新的 Google Sheet
 * 2. 工具 > Apps Script
 * 3. 把這段程式碼貼到 Code.gs
 * 4. 部署 > 新增部署 > Web App
 *    - 執行身分：自己
 *    - 存取權限：任何人
 * 5. 複製部署 URL，貼到網站的 Google Apps Script URL 欄位
 */

const SHEET_NAME = 'ETF_Data';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Store as JSON in cell A1
    sheet.getRange('A1').setValue(JSON.stringify(data));
    sheet.getRange('A2').setValue(new Date().toISOString());

    // Also write human-readable table starting from row 5
    writeReadableTable(sheet, data);

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ settings: null, entries: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const raw = sheet.getRange('A1').getValue();
    const data = raw ? JSON.parse(raw) : { settings: null, entries: [] };

    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function writeReadableTable(sheet, data) {
  // Clear old table
  if (sheet.getLastRow() > 4) {
    sheet.getRange(4, 1, sheet.getLastRow() - 3, 12).clear();
  }

  // Settings header
  sheet.getRange('A4').setValue('最後同步');
  sheet.getRange('B4').setValue(new Date().toLocaleString('zh-TW'));

  // Column headers
  const headers = ['月', '購買日期', '股價', '實際股數', '存股路徑', '累計市值', '應買股數', '投資金額', '投資累計', '股數累計', '庫存現值', '投報率'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  if (!data.entries || data.entries.length === 0) return;

  // Compute and write rows
  const s = data.settings;
  const R = (s.monthlyReturn + s.monthlyInflation) / 2;
  const T = s.months;
  const C = s.target / (T * Math.pow(1 + R, T));

  const rows = [];
  let prevCumShares = 0, prevCumInvest = 0;

  for (const e of data.entries) {
    const path = C * e.month * Math.pow(1 + R, e.month);
    const price = e.price || 0;
    const shares = e.shares || 0;
    const marketBefore = prevCumShares * price;
    const suggested = price > 0 ? (path - marketBefore) / price : 0;
    const invest = price * shares;
    const cumInvest = prevCumInvest + invest;
    const cumShares = prevCumShares + shares;
    const currentVal = price * cumShares;
    const ret = currentVal > 0 && cumInvest > 0 ? (currentVal - cumInvest) / cumInvest : 0;

    rows.push([
      e.month, e.date || '', price, shares,
      Math.round(path), Math.round(marketBefore), Math.round(suggested),
      Math.round(invest), Math.round(cumInvest), cumShares,
      Math.round(currentVal), (ret * 100).toFixed(2) + '%'
    ]);

    prevCumShares = cumShares;
    prevCumInvest = cumInvest;
  }

  sheet.getRange(6, 1, rows.length, rows[0].length).setValues(rows);
}
