/**
 * PARABELLUM — מודול הזמנות (Orders Module)
 * Backend: Google Apps Script Web App bound to a dedicated Google Sheet.
 *
 * SETUP:
 * 1. Create a new Google Sheet. Rename the first tab to exactly: Orders
 * 2. In row 1, paste these headers (in this exact order), one per column A→X:
 *    id | תאריך הזמנה | מס' לוג | צ' | צ' עומד | מקט | תיאור | סוג כלי | מודל | גרסא |
 *    סיבה להזמנה | כמות שהוזמנה | כמות שאושרה | חלק מיוחד | האם בתמורה |
 *    סטאטוס הזמנה | סיבת דחייה | כמות שנמשכה | סטאטוס בלאי | הוזמן עי | הערות |
 *    עודכן | לוג סטאטוס | לוג בלאי
 * 3. Extensions → Apps Script. Delete any starter code, paste this file's contents.
 * 4. Deploy → New deployment → Web app.
 *    Execute as: Me.  Who has access: Anyone.
 * 5. Copy the Web App URL and paste it into SCRIPT_URL in orders.html
 *
 * NOTE: if you already had this Sheet set up before, just add one new header
 * in column X: לוג בלאי — existing rows will just start with an empty log.
 */

const SHEET_NAME = 'Orders';

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'get') return getOrders();
    if (action === 'debug_headers') return debugHeaders();
    return jsonResponse({ error: 'פעולה לא מוכרת' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const action = e.parameter.action;
    const data = JSON.parse(e.parameter.data || '{}');
    if (action === 'add') return addOrder(data);
    if (action === 'update') return updateOrder(data);
    if (action === 'delete') return deleteOrder(data);
    return jsonResponse({ error: 'פעולה לא מוכרת' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('לא נמצא טאב בשם Orders');
  return sheet;
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getOrders() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse({ orders: [] });
  const headers = values[0];
  const orders = values.slice(1)
    .filter(row => row[0]) // must have an id
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  return jsonResponse({ orders: orders });
}

function formatLogDate_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jerusalem', 'dd/MM/yyyy');
}

// These fields can look like pure numbers (e.g. "000000005") — without this,
// Sheets auto-converts them to actual numbers and silently drops leading zeros.
// Setting the cell's number format to plain text BEFORE writing is the reliable
// fix — a leading apostrophe alone isn't consistently honored via the API.
const TEXT_FORCE_FIELDS = ['מקט', "צ'", "מס' לוג"];

function findColIndex_(headers, fieldName) {
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === fieldName) return i;
  }
  return -1;
}

function preFormatTextColumns_(sheet, rowIndex, headers) {
  TEXT_FORCE_FIELDS.forEach(fieldName => {
    const colIdx = findColIndex_(headers, fieldName);
    if (colIdx !== -1) {
      sheet.getRange(rowIndex, colIdx + 1).setNumberFormat('@');
    }
  });
  SpreadsheetApp.flush(); // force the format change to actually commit before we write values
}

// TEMPORARY — hit ?action=debug_headers in the browser to see exactly what
// Apps Script sees in row 1, and whether it matches TEXT_FORCE_FIELDS.
function debugHeaders() {
  const sheet = getSheet_();
  const headers = getHeaders_(sheet);
  const check = TEXT_FORCE_FIELDS.map(f => ({
    field: f,
    foundAt: findColIndex_(headers, f),
  }));
  return jsonResponse({ headers: headers, textForceCheck: check });
}

function addOrder(data) {
  const sheet = getSheet_();
  const headers = getHeaders_(sheet);
  const id = 'ORD-' + new Date().getTime();
  data.id = id;
  data['עודכן'] = new Date().toISOString();
  const who = data['עודכן עי'] || data['הוזמן עי'] || 'לא ידוע';
  const initialStatus = data['סטאטוס הזמנה'] || 'נשלח';
  data['לוג סטאטוס'] = `${formatLogDate_()}|${initialStatus}|${who}`;
  if (data['סטאטוס בלאי']) {
    data['לוג בלאי'] = `${formatLogDate_()}|${data['סטאטוס בלאי']}|${who}`;
  }
  const row = headers.map(h => (data[h] !== undefined ? data[h] : ''));
  const newRowIndex = sheet.getLastRow() + 1;
  preFormatTextColumns_(sheet, newRowIndex, headers); // format BEFORE writing values
  sheet.getRange(newRowIndex, 1, 1, row.length).setValues([row]);
  return jsonResponse({ success: true, id: id });
}

function updateOrder(data) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  const statusCol = headers.indexOf('סטאטוס הזמנה');
  const logCol = headers.indexOf('לוג סטאטוס');
  const blaiStatusCol = headers.indexOf('סטאטוס בלאי');
  const blaiLogCol = headers.indexOf('לוג בלאי');
  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === data.id) {
      data['עודכן'] = new Date().toISOString();

      const oldStatus = values[i][statusCol];
      const newStatus = data['סטאטוס הזמנה'];
      if (newStatus !== undefined && newStatus !== oldStatus && logCol !== -1) {
        const who = data['עודכן עי'] || 'לא ידוע';
        const existingLog = values[i][logCol] || '';
        const entry = `${formatLogDate_()}|${newStatus}|${who}`;
        data['לוג סטאטוס'] = existingLog ? existingLog + '\n' + entry : entry;
      }

      const oldBlaiStatus = blaiStatusCol !== -1 ? values[i][blaiStatusCol] : undefined;
      const newBlaiStatus = data['סטאטוס בלאי'];
      if (newBlaiStatus !== undefined && newBlaiStatus !== oldBlaiStatus && blaiLogCol !== -1) {
        const who = data['עודכן עי'] || 'לא ידוע';
        const existingBlaiLog = values[i][blaiLogCol] || '';
        const entry = `${formatLogDate_()}|${newBlaiStatus}|${who}`;
        data['לוג בלאי'] = existingBlaiLog ? existingBlaiLog + '\n' + entry : entry;
      }

      const rowIndex = i + 1;
      preFormatTextColumns_(sheet, rowIndex, headers); // format BEFORE writing values
      const row = headers.map((h, colIdx) =>
        data[h] !== undefined ? data[h] : values[i][colIdx]
      );
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ error: 'הזמנה לא נמצאה' });
}

function deleteOrder(data) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === data.id) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ error: 'הזמנה לא נמצאה' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
