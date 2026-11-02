/**
 * Villa Expense Ledger — Google Sheet backend.
 *
 * Paste this file into the Google Sheet's Apps Script editor
 * (Extensions → Apps Script), set the two PINs below, then
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * The app talks to the resulting ".../exec" link. Every entry is a row
 * in the "Entries" tab. After each change the script also rebuilds that
 * month's tab ("Nov 2026": day-wise, with totals and a filter) and the
 * "Summary" tab. Run rebuildAllMonths() once from the editor to rebuild
 * every month tab from "Entries".
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
      if (req.action === 'save') {
        const saved = save_(req.entry || {});
        refreshMonths_([saved.date, saved.previousDate]);
        delete saved.previousDate;
        return json_({ ok: true, entry: saved });
      }
      if (req.action === 'delete') {
        refreshMonths_([delete_(String(req.id || ''))]);
        return json_({ ok: true });
      }
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
  const previousDate = row ? dateText_(sh.getRange(row, 2).getValue()) : '';
  const values = [[id, date, type === 'credit' ? 'Credit' : 'Debit', CATEGORIES[category], description,
    String(e.person || '').trim().slice(0, 80), method, amount, String(e.note || '').trim().slice(0, 400), created, now]];
  if (row) sh.getRange(row, 1, 1, HEADERS.length).setValues(values);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, HEADERS.length).setValues(values);
  return { id: id, date: date, type: type, category: category, description: values[0][4], person: values[0][5],
    method: method, amount: amount, note: values[0][8], createdAt: (created instanceof Date ? created : now).getTime(),
    previousDate: previousDate };
}

function delete_(id) {
  if (!id) throw new Error('Missing entry id.');
  const sh = sheet_();
  const row = findRow_(sh, id);
  if (!row) return '';
  const date = dateText_(sh.getRange(row, 2).getValue());
  sh.deleteRow(row);
  return date;
}

function dateText_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v || '');
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}

/* ---------- Month tabs and Summary ---------- */

const MONTH_CATS = ['Utilities', 'Visa', 'Tickets', 'Villa cash'];
const HEAD_FILL = '#e0f0f2';

// Rebuild every month tab and the Summary from "Entries". Run it from the editor.
function rebuildAllMonths() {
  const months = {};
  list_().forEach(function (e) { months[e.date.slice(0, 7)] = true; });
  refreshMonths_(Object.keys(months).map(function (m) { return m + '-01'; }));
}

function refreshMonths_(dates) {
  const all = list_();
  const done = {};
  dates.forEach(function (d) {
    const ym = String(d || '').slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(ym) || done[ym]) return;
    done[ym] = true;
    writeMonth_(ym, all.filter(function (e) { return e.date.indexOf(ym) === 0; }));
  });
  writeSummary_(all);
}

function monthTitle_(ym) {
  const p = ym.split('-');
  return Utilities.formatDate(new Date(Number(p[0]), Number(p[1]) - 1, 1, 12), Session.getScriptTimeZone(), 'MMM yyyy');
}

function tab_(title, index) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(title);
  if (!sh) sh = ss.insertSheet(title, Math.min(index, ss.getNumSheets()));
  const filter = sh.getFilter();
  if (filter) filter.remove();
  sh.clear();
  return sh;
}

// Text that starts like a formula is kept as plain text.
function text_(v) {
  const t = String(v == null ? '' : v);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function writeMonth_(ym, entries) {
  entries.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.createdAt || 0) - (b.createdAt || 0); });
  const sh = tab_(monthTitle_(ym), 1);
  // Totals sit above the header so the filter never hides them.
  const top = [
    ['Credit', '=SUM(H5:H)', 'Debit', '=SUM(I5:I)', 'Balance', '=B1-D1', 'Shown by filter:', 'Credit', '=SUBTOTAL(109,H5:H)', 'Debit', '=SUBTOTAL(109,I5:I)'],
    [].concat.apply([], MONTH_CATS.map(function (c) { return [c, '=SUMIFS(I5:I,D5:D,"' + c + '")']; })).concat(['', '', '']),
    ['', '', '', '', '', '', '', '', '', '', ''],
  ];
  sh.getRange(1, 1, top.length, 11).setValues(top);
  sh.getRange(1, 1, 2, 11).setFontWeight('bold').setNumberFormat('#,##0.00');

  const header = ['Date', 'Day', 'Type', 'Category', 'Description', 'Person', 'Paid by', 'Credit (AED)', 'Debit (AED)', 'Note'];
  const rows = entries.map(function (e) {
    const p = e.date.split('-').map(Number);
    const day = new Date(p[0], p[1] - 1, p[2], 12);
    const credit = e.type === 'credit';
    return [day, Utilities.formatDate(day, Session.getScriptTimeZone(), 'EEEE'), credit ? 'Credit' : 'Debit',
      CATEGORIES[e.category] || e.category, text_(e.description), text_(e.person), text_(e.method),
      credit ? e.amount : '', credit ? '' : e.amount, text_(e.note)];
  });
  const table = sh.getRange(4, 1, rows.length + 1, header.length);
  table.setValues([header].concat(rows));
  sh.getRange(4, 1, 1, header.length).setFontWeight('bold').setBackground(HEAD_FILL);
  if (rows.length) {
    sh.getRange(5, 1, rows.length, 1).setNumberFormat('dd-mmm-yyyy');
    sh.getRange(5, 8, rows.length, 2).setNumberFormat('#,##0.00');
  }
  table.createFilter();
  sh.setFrozenRows(4);
  sh.autoResizeColumns(1, 11);
}

function writeSummary_(entries) {
  const sh = tab_('Summary', 0);
  const byMonth = {};
  entries.forEach(function (e) {
    const ym = e.date.slice(0, 7);
    const m = byMonth[ym] || (byMonth[ym] = { credit: 0, debit: 0, cats: {} });
    if (e.type === 'credit') m.credit += e.amount;
    else { m.debit += e.amount; m.cats[e.category] = (m.cats[e.category] || 0) + e.amount; }
  });
  const keys = ['utilities', 'visa', 'tickets', 'villa'];
  const rows = Object.keys(byMonth).sort().reverse().map(function (ym) {
    const m = byMonth[ym];
    return [monthTitle_(ym), m.credit, m.debit, m.credit - m.debit].concat(keys.map(function (k) { return m.cats[k] || 0; }));
  });
  const header = ['Month', 'Credit', 'Debit', 'Balance'].concat(MONTH_CATS);
  sh.getRange(1, 1, rows.length + 1, header.length).setValues([header].concat(rows));
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground(HEAD_FILL);
  if (rows.length) sh.getRange(2, 2, rows.length, header.length - 1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
