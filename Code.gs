/**
 * PARABELLUM — מודול הזמנות (Orders Module)
 * Backend: Google Apps Script Web App bound to a dedicated Google Sheet.
 *
 * SETUP:
 * 1. Create a new Google Sheet. Rename the first tab to exactly: Orders
 * 2. In row 1, paste these headers (in this exact order), one per column A→V:
 *    id | תאריך הזמנה | מס' לוג | צ' | צ' עומד | מקט | תיאור | סוג כלי | מודל | גרסא |
 *    סיבה להזמנה | כמות שהוזמנה | כמות שאושרה | חלק מיוחד | האם בתמורה |
 *    סטאטוס הזמנה | סיבת דחייה | כמות שנמשכה | סטאטוס בלאי | הוזמן עי | הערות | עודכן
 * 3. Extensions → Apps Script. Delete any starter code, paste this file's contents.
 * 4. Deploy → New deployment → Web app.
 *    Execute as: Me.  Who has access: Anyone.
 * 5. Copy the Web App URL and paste it into SCRIPT_URL in orders.html
 */

const SHEET_NAME = 'Orders';

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'get') return getOrders();
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

function addOrder(data) {
  const sheet = getSheet_();
  const headers = getHeaders_(sheet);
  const id = 'ORD-' + new Date().getTime();
  data.id = id;
  data['עודכן'] = new Date().toISOString();
  const row = headers.map(h => (data[h] !== undefined ? data[h] : ''));
  sheet.appendRow(row);
  return jsonResponse({ success: true, id: id });
}

function updateOrder(data) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === data.id) {
      data['עודכן'] = new Date().toISOString();
      const row = headers.map((h, colIdx) =>
        data[h] !== undefined ? data[h] : values[i][colIdx]
      );
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
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
