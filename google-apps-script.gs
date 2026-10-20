/**
 * ledger backend (personal Google account only).
 * Stores entries in this Google Sheet and receipt photos in Google Drive
 * under "ledger receipts / YEAR / MONTH".
 *
 * Setup: see SETUP.md. Change SECRET to your own passcode before deploying.
 */
const SECRET = 'CHANGE-ME';

// version 2 (2026-10-20): first chair reimbursements: reimb, submittedOn, paidOn, trip
// version 3 (2026-10-20): 'by category' tab listing every expense under its category
const VERSION = 3;
const HEAD = ['id', 'date', 'type', 'amount', 'category', 'merchant', 'scope', 'note', 'source', 'receiptId', 'receiptUrl', 'created', 'reimb', 'submittedOn', 'paidOn', 'trip'];
const ROOT_FOLDER = 'ledger receipts';

function doGet() {
  return ContentService.createTextOutput('ledger is running.');
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad request' }); }
  if (!body || body.key !== SECRET) return out_({ ok: false, error: 'wrong passcode' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (body.action) {
      case 'ping': return out_({ ok: true, version: VERSION });
      case 'list': return out_({ ok: true, version: VERSION, entries: list_(), settings: getSettings_() });
      case 'add':
      case 'update': { const entry = upsert_(body.entry, body.receipt); refreshCategories_(); return out_({ ok: true, entry }); }
      case 'delete': remove_(body.entry.id); refreshCategories_(); return out_({ ok: true });
      case 'settings': saveSettings_(body.settings); refreshCategories_(); return out_({ ok: true });
      case 'receipt': return out_(Object.assign({ ok: true }, receipt_(body.id)));
      case 'taxExport': return out_(Object.assign({ ok: true }, taxExport_(Number(body.year))));
      default: return out_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- sheets ----------

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  return sh;
}

function entries_() {
  const sh = sheet_('entries', HEAD);
  // add any new columns to sheets made by an older version
  const have = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  if (have.length < HEAD.length || HEAD.some((h, i) => have[i] !== h)) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  }
  ['B:B', 'N:N', 'O:O'].forEach(c => sh.getRange(c).setNumberFormat('@')); // keep dates as plain yyyy-mm-dd text
  return sh;
}

function list_() {
  const sh = entries_();
  const rows = sh.getDataRange().getValues().slice(1);
  const tz = Session.getScriptTimeZone();
  return rows.filter(r => r[0]).map(r => {
    const o = {};
    HEAD.forEach((h, i) => {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, h === 'created' ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
      if (v === undefined) v = '';
      o[h] = v;
    });
    o.amount = Number(o.amount) || 0;
    return o;
  });
}

function rowOf_(e) {
  return HEAD.map(h => h === 'amount' ? Number(e[h]) || 0 : (e[h] == null ? '' : String(e[h])));
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function upsert_(e, receiptDataUrl) {
  if (!e || !e.id) throw new Error('missing entry');
  const sh = entries_();
  const r = findRow_(sh, e.id);
  if (receiptDataUrl) {
    if (e.receiptId) trash_(e.receiptId);
    const f = saveReceipt_(e, receiptDataUrl);
    e.receiptId = f.getId();
    e.receiptUrl = f.getUrl();
  }
  if (r) sh.getRange(r, 1, 1, HEAD.length).setValues([rowOf_(e)]);
  else sh.appendRow(rowOf_(e));
  return e;
}

function remove_(id) {
  const sh = entries_();
  const r = findRow_(sh, id);
  if (!r) return;
  const receiptId = sh.getRange(r, HEAD.indexOf('receiptId') + 1).getValue();
  if (receiptId) trash_(receiptId);
  sh.deleteRow(r);
}

function getSettings_() {
  const sh = sheet_('settings', ['settings (managed by the app, do not edit)']);
  const v = sh.getRange(2, 1).getValue();
  try { return v ? JSON.parse(v) : null; } catch (err) { return null; }
}

function saveSettings_(s) {
  const sh = sheet_('settings', ['settings (managed by the app, do not edit)']);
  sh.getRange(2, 1).setValue(JSON.stringify(s || {}));
}

// ---------- 'by category' tab ----------
// rebuilt after every change: each month, each category with its subtotal, every expense under it

const MONTH_NAMES = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
const WHO = { personal: 'personal', business: 'business', firstchair: 'first chair' };

function refreshCategories_() {
  try { buildCategories_(); } catch (err) { console.error('by category tab: ' + err); }
}

function buildCategories_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('by category') || ss.insertSheet('by category', 1);
  const entries = list_().filter(e => e.type === 'expense');
  const settings = getSettings_() || {};
  const bills = settings.bills || [];
  const tz = Session.getScriptTimeZone();
  const thisMonth = Utilities.formatDate(new Date(), tz, 'yyyy-MM');

  const months = Array.from(new Set(entries.map(e => String(e.date).slice(0, 7)).concat(thisMonth))).filter(Boolean).sort().reverse();
  const W = 6, rows = [['category', 'date', 'store', 'amount', 'who', 'note']], monthRows = [], catRows = [];
  const push = r => rows.push(r.concat(Array(W - r.length).fill('')));

  months.forEach(m => {
    const items = entries.filter(e => String(e.date).slice(0, 7) === m).map(e => ({
      date: e.date, store: e.merchant, amount: Number(e.amount) || 0, category: e.category || 'other', scope: e.scope || 'personal',
      note: [e.note, e.scope === 'firstchair' ? ({ '': 'to submit', submitted: 'submitted', paid: 'paid back' })[e.reimb || ''] : ''].filter(String).join(' · '),
    }));
    // fixed bills count every month (from the first month you used the app)
    if (m >= months[months.length - 1]) bills.forEach(b => items.push({
      date: m + '-' + String(b.day || 1).padStart(2, '0'), store: b.name, amount: Number(b.amount) || 0,
      category: b.category || 'other', scope: b.scope || 'personal', note: 'fixed bill',
    }));
    const own = items.filter(i => i.scope !== 'firstchair'), fc = items.filter(i => i.scope === 'firstchair');
    const sum = a => a.reduce((t, i) => t + i.amount, 0);
    const [y, mo] = m.split('-');

    push(['']);
    monthRows.push(rows.length + 1);
    push([MONTH_NAMES[Number(mo) - 1] + ' ' + y + '  ·  your spending', '', '', sum(own)]);
    const groups = {};
    own.forEach(i => (groups[i.category] = groups[i.category] || []).push(i));
    Object.keys(groups).sort((a, b) => sum(groups[b]) - sum(groups[a])).forEach(cat => {
      catRows.push(rows.length + 1);
      push([cat, '', groups[cat].length + (groups[cat].length === 1 ? ' expense' : ' expenses'), sum(groups[cat])]);
      groups[cat].sort((a, b) => String(a.date).localeCompare(String(b.date))).forEach(i => push(['', i.date, i.store, i.amount, WHO[i.scope] || i.scope, i.note]));
    });
    if (!own.length) push(['', '', 'nothing logged']);
    if (fc.length) {
      catRows.push(rows.length + 1);
      push(['first chair (paid back to you)', '', fc.length + (fc.length === 1 ? ' charge' : ' charges'), sum(fc)]);
      fc.sort((a, b) => String(a.date).localeCompare(String(b.date))).forEach(i => push(['', i.date, i.store, i.amount, 'first chair', i.note]));
    }
  });

  sh.clear();
  sh.getRange(1, 1, rows.length, W).setValues(rows);
  sh.getRange(1, 1, 1, W).setFontWeight('bold').setBackground('#212121').setFontColor('#ffffff');
  sh.getRange(2, 2, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 4, rows.length, 1).setNumberFormat('$#,##0.00');
  monthRows.forEach(r => sh.getRange(r, 1, 1, W).setFontWeight('bold').setFontSize(12).setBackground('#ebe8e1'));
  catRows.forEach(r => sh.getRange(r, 1, 1, W).setFontWeight('bold').setBackground('#f6f4f0'));
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 230); sh.setColumnWidth(2, 95); sh.setColumnWidth(3, 200); sh.setColumnWidth(4, 95); sh.setColumnWidth(5, 95); sh.setColumnWidth(6, 200);
}

// ---------- receipts in Drive ----------

function folder_(parent, name) {
  const it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

function yearFolder_(year) {
  return folder_(folder_(null, ROOT_FOLDER), String(year));
}

function saveReceipt_(e, dataUrl) {
  const m = String(dataUrl).match(/^data:([^;]+);base64,(.*)$/);
  if (!m) throw new Error('bad receipt image');
  const date = String(e.date || '').slice(0, 10);
  const [y, mo] = date.split('-');
  const folder = folder_(yearFolder_(y || new Date().getFullYear()), mo || '00');
  const who = String(e.merchant || e.category || 'receipt').toLowerCase().replace(/[^a-z0-9 ]+/g, '').trim().slice(0, 40) || 'receipt';
  const ext = m[1] === 'image/png' ? 'png' : 'jpg';
  const name = `${date} ${who} ${Number(e.amount || 0).toFixed(2)}.${ext}`;
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
  return folder.createFile(blob);
}

function trash_(id) {
  try { DriveApp.getFileById(id).setTrashed(true); } catch (err) { /* already gone */ }
}

function receipt_(id) {
  const f = DriveApp.getFileById(id);
  const b = f.getBlob();
  return { mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()), url: f.getUrl(), name: f.getName() };
}

// ---------- year-end tax export ----------

function taxExport_(year) {
  if (!year) throw new Error('missing year');
  const all = list_().filter(e => String(e.date).slice(0, 4) === String(year));
  const settings = getSettings_() || {};
  const bills = settings.bills || [];
  const paid = settings.paid || {};

  const rows = all
    .filter(e => e.type === 'expense' && e.scope === 'business')
    .map(e => [e.date, e.merchant, e.category, e.amount, e.note, e.receiptUrl ? e.receiptUrl : 'no receipt']);

  // business fixed bills that were ticked paid, one row per month
  Object.keys(paid).filter(k => k.slice(0, 4) === String(year)).sort().forEach(month => {
    (paid[month] || []).forEach(billId => {
      const b = bills.find(x => x.id === billId);
      if (b && b.scope === 'business') rows.push([`${month}-${String(b.day || 1).padStart(2, '0')}`, b.name, b.category, Number(b.amount) || 0, 'fixed bill', 'fixed bill']);
    });
  });
  rows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  const income = all.filter(e => e.type === 'income')
    .map(e => [e.date, e.merchant, e.category, e.scope, e.amount, e.note])
    .sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  const byCat = {};
  rows.forEach(r => { byCat[r[2] || 'other'] = (byCat[r[2] || 'other'] || 0) + Number(r[3]); });

  const name = `ledger tax export ${year}`;
  const old = yearFolder_(year).getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);

  const ss = SpreadsheetApp.create(name);
  const s1 = ss.getSheets()[0];
  s1.setName('business expenses');
  s1.appendRow(['date', 'store', 'category', 'amount', 'note', 'receipt']);
  if (rows.length) s1.getRange(2, 1, rows.length, 6).setValues(rows);
  s1.appendRow(['', '', 'total', rows.reduce((t, r) => t + Number(r[3]), 0)]);

  const s2 = ss.insertSheet('business by category');
  s2.appendRow(['category', 'total']);
  const cats = Object.keys(byCat).sort((a, b) => byCat[b] - byCat[a]).map(k => [k, byCat[k]]);
  if (cats.length) s2.getRange(2, 1, cats.length, 2).setValues(cats);

  const s3 = ss.insertSheet('income');
  s3.appendRow(['date', 'from', 'category', 'personal or business', 'amount', 'note']);
  if (income.length) s3.getRange(2, 1, income.length, 6).setValues(income);
  s3.appendRow(['', '', '', 'total', income.reduce((t, r) => t + Number(r[4]), 0)]);

  [s1, s2, s3].forEach(s => { s.setFrozenRows(1); s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold'); });
  s1.getRange('D:D').setNumberFormat('$#,##0.00');
  s2.getRange('B:B').setNumberFormat('$#,##0.00');
  s3.getRange('E:E').setNumberFormat('$#,##0.00');

  const folder = yearFolder_(year);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  return { sheetUrl: ss.getUrl(), folderUrl: folder.getUrl(), count: rows.length };
}
