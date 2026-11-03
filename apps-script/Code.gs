/**
 * Villa Expense Ledger — Google Sheet backend.
 *
 * Paste this file into the Google Sheet's Apps Script editor
 * (Extensions → Apps Script), set the two PINs below, then
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * The app talks to the resulting ".../exec" link. Every entry is a row
 * in the "Entries" tab, shared with the Claude ledger page. Month tabs
 * ("Nov 2026": day-wise, totals and a filter) and the "Summary" tab are
 * formulas over "Entries"; the script adds a month's tab the first time that
 * month appears. Run rebuildAllMonths() from the editor to add any missing.
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

/* ---------- Month tabs and Summary ----------
   Month tabs ("Nov 2026") and the Summary are formulas over "Entries", so they
   stay correct whichever device saved an entry (this app or the Claude
   ledger). A tab is only written when it is missing or still holds an old
   fixed copy (no marker in M1). */

const MONTH_CATS = ['Utilities', 'Visa', 'Tickets', 'Villa cash'];
const HEAD_FILL = '#e0f0f2';
const MARK = 'Updates by itself from the Entries tab. Add or change entries in the app or the Claude ledger, not here.';
const KEY = 'IF(ISNUMBER(Entries!B2:B),TEXT(Entries!B2:B,"yyyy-mm-dd"),Entries!B2:B)';

// Make sure every month in "Entries" has its tab and Summary row. Run it from the editor.
function rebuildAllMonths() {
  refreshMonths_([]);
}

function refreshMonths_(dates) {
  const months = {};
  list_().forEach(function (e) { months[e.date.slice(0, 7)] = true; });
  dates.forEach(function (d) { const ym = String(d || '').slice(0, 7); if (/^\d{4}-\d{2}$/.test(ym)) months[ym] = true; });
  const list = Object.keys(months).sort().reverse();
  list.forEach(ensureMonth_);
  ensureSummary_(list);
}

function monthTitle_(ym) {
  const p = ym.split('-');
  return Utilities.formatDate(new Date(Number(p[0]), Number(p[1]) - 1, 1, 12), Session.getScriptTimeZone(), 'MMM yyyy');
}

function ensureMonth_(ym) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const title = monthTitle_(ym);
  let sh = ss.getSheetByName(title);
  if (sh && sh.getRange('M1').getValue() === MARK) return;
  if (!sh) sh = ss.insertSheet(title, Math.min(1, ss.getNumSheets()));
  const filter = sh.getFilter();
  if (filter) filter.remove();
  sh.clear();
  const rows = '=ARRAYFORMULA(IFERROR(SORT(FILTER({DATEVALUE(' + KEY + '),TEXT(DATEVALUE(' + KEY + '),"dddd"),Entries!C2:C,Entries!D2:D,' +
    'Entries!E2:E,Entries!F2:F,Entries!G2:G,IF(Entries!C2:C="Credit",Entries!H2:H,""),IF(Entries!C2:C="Debit",Entries!H2:H,""),Entries!I2:I},' +
    'LEFT(' + KEY + ',7)="' + ym + '"),1,TRUE),""))';
  const pad = function (r) { while (r.length < 13) r.push(''); return r; };
  sh.getRange(1, 1, 5, 13).setValues([
    pad(['Credit', '=SUM(H5:H)', 'Debit', '=SUM(I5:I)', 'Balance', '=B1-D1', 'Shown by filter:', 'Credit', '=SUBTOTAL(109,H5:H)', 'Debit', '=SUBTOTAL(109,I5:I)', '', MARK]),
    pad([].concat.apply([], MONTH_CATS.map(function (c) { return [c, '=SUMIFS(I5:I,D5:D,"' + c + '")']; }))),
    pad([]),
    pad(['Date', 'Day', 'Type', 'Category', 'Description', 'Person', 'Paid by', 'Credit (AED)', 'Debit (AED)', 'Note']),
    pad([rows]),
  ]);
  sh.getRange(1, 1, 2, 11).setFontWeight('bold').setNumberFormat('#,##0.00');
  sh.getRange('M1').setFontStyle('italic');
  sh.getRange(4, 1, 1, 10).setFontWeight('bold').setBackground(HEAD_FILL);
  sh.getRange('A5:A').setNumberFormat('dd-mmm-yyyy');
  sh.getRange('H5:I').setNumberFormat('#,##0.00');
  sh.getRange(4, 1, sh.getMaxRows() - 3, 10).createFilter();
  sh.setFrozenRows(4);
  sh.autoResizeColumns(1, 11);
}

function ensureSummary_(months) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Summary');
  if (!sh) sh = ss.insertSheet('Summary', 0);
  const want = months.map(monthTitle_);
  const n = Math.max(sh.getLastRow() - 1, 0);
  const have = n ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  if (sh.getRange('A1').getValue() === 'Month' && have.join('|') === want.join('|')) return;
  const header = ['Month', 'Credit', 'Debit', 'Balance'].concat(MONTH_CATS);
  const rows = months.map(function (ym, i) {
    const r = i + 2, m = '"' + ym + '*"';
    return [monthTitle_(ym), '=SUMIFS(Entries!H:H,Entries!B:B,' + m + ',Entries!C:C,"Credit")',
      '=SUMIFS(Entries!H:H,Entries!B:B,' + m + ',Entries!C:C,"Debit")', '=B' + r + '-C' + r]
      .concat(MONTH_CATS.map(function (c) { return '=SUMIFS(Entries!H:H,Entries!B:B,' + m + ',Entries!C:C,"Debit",Entries!D:D,"' + c + '")'; }));
  });
  sh.clear();
  sh.getRange(1, 1, rows.length + 1, header.length).setValues([header].concat(rows));
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground(HEAD_FILL);
  if (rows.length) sh.getRange(2, 2, rows.length, header.length - 1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
