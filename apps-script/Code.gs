/**
 * Villa Expense Ledger — Google Sheet backend.
 *
 * Paste this file into the Google Sheet's Apps Script editor
 * (Extensions → Apps Script), set the two PINs below, then
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * The app talks to the resulting ".../exec" link. Every entry is a row
 * in the "Entries" tab, so the sheet itself is also a readable record.
 */

// PIN for adding, editing and deleting entries. Change it before deploying.
const EDIT_PIN = 'CHANGE-ME';
// PIN for view-only access (e.g. for your boss). Leave '' to turn it off.
const VIEW_PIN = '';

const SHEET_NAME = 'Entries';
const HEADERS = ['ID', 'Date', 'Type', 'Category', 'Description', 'Person', 'Paid by', 'Amount', 'Note', 'Created', 'Updated'];
const CATEGORIES = { utilities: 'Utilities', visa: 'Visa', tickets: 'Tickets', villa: 'Villa cash', received: 'Received' };
const METHODS = ['Cash', 'Bank transfer', 'Card', 'Other'];

function doGet() {
  return json_({ ok: true, app: 'villa-expense-ledger', message: 'Backend is running. Open the app and paste this link into its setup screen.' });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  try {
    const role = roleFor_(req.pin);
    if (!role) return json_({ ok: false, error: 'wrong_pin' });
    if (req.action === 'list') return json_({ ok: true, role: role, entries: list_() });
    if (role !== 'edit') return json_({ ok: false, error: 'read_only' });

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (req.action === 'save') return json_({ ok: true, entry: save_(req.entry || {}) });
      if (req.action === 'delete') { delete_(String(req.id || '')); return json_({ ok: true }); }
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server_error', message: String(err && err.message || err) });
  }
}

function roleFor_(pin) {
  if (!pin) return null;
  if (EDIT_PIN && EDIT_PIN !== 'CHANGE-ME' && pin === EDIT_PIN) return 'edit';
  if (VIEW_PIN && pin === VIEW_PIN) return 'view';
  return null;
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // Keep dates as plain text so Sheets doesn't reformat them.
  sh.getRange('B:B').setNumberFormat('@');
  return sh;
}

function list_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const tz = Session.getScriptTimeZone();
  const keyByName = {};
  Object.keys(CATEGORIES).forEach(function (k) { keyByName[CATEGORIES[k]] = k; });
  return sh.getRange(2, 1, n, HEADERS.length).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      const date = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]);
      return {
        id: String(r[0]), date: date, type: String(r[2]).toLowerCase() === 'credit' ? 'credit' : 'debit',
        category: keyByName[r[3]] || String(r[3]), description: String(r[4]), person: String(r[5]),
        method: String(r[6]), amount: Number(r[7]), note: String(r[8]),
        createdAt: r[9] instanceof Date ? r[9].getTime() : Number(r[9]) || 0,
      };
    });
}

function save_(e) {
  const type = e.type === 'credit' ? 'credit' : 'debit';
  const category = type === 'credit' ? 'received' : String(e.category || '');
  const amount = Math.round(Number(e.amount) * 100) / 100;
  const date = String(e.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Enter a valid date.');
  if (!CATEGORIES[category]) throw new Error('Choose Utilities, Visa, Tickets or Villa cash.');
  if (!(amount > 0)) throw new Error('Enter an amount greater than zero.');
  const description = String(e.description || '').trim().slice(0, 140);
  if (!description) throw new Error('Add a short description.');
  const method = METHODS.indexOf(e.method) >= 0 ? e.method : 'Cash';

  const sh = sheet_();
  const now = new Date();
  const id = e.id ? String(e.id) : Utilities.getUuid();
  const row = findRow_(sh, id);
  const created = row ? sh.getRange(row, 10).getValue() : now;
  const values = [[id, date, type === 'credit' ? 'Credit' : 'Debit', CATEGORIES[category], description,
    String(e.person || '').trim().slice(0, 80), method, amount, String(e.note || '').trim().slice(0, 400), created, now]];
  if (row) sh.getRange(row, 1, 1, HEADERS.length).setValues(values);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, HEADERS.length).setValues(values);
  return { id: id, date: date, type: type, category: category, description: values[0][4], person: values[0][5],
    method: method, amount: amount, note: values[0][8], createdAt: (created instanceof Date ? created : now).getTime() };
}

function delete_(id) {
  if (!id) throw new Error('Missing entry id.');
  const sh = sheet_();
  const row = findRow_(sh, id);
  if (row) sh.deleteRow(row);
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
