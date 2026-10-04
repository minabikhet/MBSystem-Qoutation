/**
 * MB System – Quotations sheet
 * Every quote saved from the app is written to its own tab (named after the client)
 * and listed in the "العروض" tab. The same script also serves the app itself (file Index.html).
 */
const INDEX = 'العروض';
const HEAD = ['ID', 'العميل', 'التاريخ', 'الإجمالي', 'العملة', 'اللغة', 'آخر حفظ', 'التاب', 'DATA'];
const GOLD = '#ae8037', NAVY = '#233548', SOFT = '#f3ead9';

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'list') return out({ ok: true, quotes: listQuotes() });
  if (action === 'ping') return out({ ok: true });
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MB System Quotations')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    ;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'save') { saveQuote(body.quote); return out({ ok: true }); }
    if (body.action === 'delete') { deleteQuote(body.id); return out({ ok: true }); }
    return out({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function indexSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(INDEX);
  if (!sh) {
    sh = ss.insertSheet(INDEX, 0);
    sh.setRightToLeft(true);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD])
      .setFontWeight('bold').setBackground(NAVY).setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.hideColumns(1);
    sh.hideColumns(9);
    sh.setColumnWidths(2, 7, 130);
    if (sh.getMaxColumns() > 9) sh.deleteColumns(10, sh.getMaxColumns() - 9);
    const first = ss.getSheets().find(s => s.getName() !== INDEX && s.getLastRow() === 0);
    if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  }
  return sh;
}

function findRow(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function listQuotes() {
  const sh = indexSheet();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 9, n - 1, 1).getValues()
    .map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } })
    .filter(Boolean);
}

function cleanName(s) {
  return (String(s || 'Quote').replace(/[\[\]\*\?\/\\:']/g, ' ').trim().slice(0, 80)) || 'Quote';
}

function uniqueName(base, keep) {
  const ss = SpreadsheetApp.getActive();
  let name = base, i = 2;
  while (ss.getSheetByName(name) && name !== keep) name = base + ' (' + (i++) + ')';
  return name;
}

function dmy(s) { if (!s) return ''; const p = String(s).split('-'); return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : s; }
function num(v) { const n = parseFloat(String(v).replace(/,/g, '')); return isNaN(n) ? 0 : n; }

function saveQuote(q) {
  const ss = SpreadsheetApp.getActive();
  const idx = indexSheet();
  let row = findRow(idx, q.id);
  const oldTab = row > 0 ? idx.getRange(row, 8).getValue() : '';
  const base = cleanName(q.client) + ' - ' + dmy(q.date).replace(/\//g, '-');
  let tabName = oldTab && oldTab.indexOf(cleanName(q.client)) === 0 ? oldTab : uniqueName(base, oldTab);

  let sh = oldTab ? ss.getSheetByName(oldTab) : null;
  if (!sh) sh = ss.insertSheet(tabName);
  else if (oldTab !== tabName) sh.setName(tabName);
  writeQuoteTab(sh, q);

  const link = '=HYPERLINK("#gid=' + sh.getSheetId() + '","' + tabName.replace(/"/g, '""') + '")';
  const vals = [q.id, q.client, dmy(q.date), Number(q.total) || 0, q.currency || '', q.lang === 'en' ? 'English' : 'عربي',
    new Date(q.savedAt || Date.now()), link, JSON.stringify(q)];
  if (row < 0) { idx.appendRow(vals); row = idx.getLastRow(); }
  else idx.getRange(row, 1, 1, vals.length).setValues([vals]);
  idx.getRange(row, 4).setNumberFormat('#,##0.00');
  idx.getRange(row, 7).setNumberFormat('dd/MM/yyyy HH:mm');
}

function writeQuoteTab(sh, q) {
  const en = q.lang === 'en';
  const L = en
    ? { title: 'QUOTATION - MB System', client: 'Client', date: 'Date', no: '#', desc: 'Description', qty: 'Qty', price: 'Unit Price', total: 'Total', notes: 'Notes' }
    : { title: 'عرض سعر - MB System', client: 'اسم العميل', date: 'التاريخ', no: 'م', desc: 'البيان والمواصفات', qty: 'الكمية', price: 'السعر', total: 'الإجمالي', notes: 'ملاحظات' };
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  const need = 30 + ((q.items || []).length) + ((q.notes || []).length);
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.setRightToLeft(!en);
  sh.getRange(1, 1, 1, 5).merge().setValue(L.title).setFontSize(16).setFontWeight('bold').setFontColor(NAVY);
  sh.getRange(3, 1, 2, 2).setValues([[L.client, q.client], [L.date, dmy(q.date)]]);
  sh.getRange(3, 1, 2, 1).setFontWeight('bold').setFontColor(GOLD);

  const head = 6;
  sh.getRange(head, 1, 1, 5).setValues([[L.no, L.desc, L.qty, L.price, L.total]])
    .setFontWeight('bold').setBackground(SOFT).setFontColor(NAVY).setHorizontalAlignment('center');
  const items = (q.items || []).filter(it => String(it.d || '').trim());
  const rows = items.map((it, i) => [i + 1, it.d, num(it.q), num(it.p), Math.round(num(it.q) * num(it.p) * 100) / 100]);
  if (rows.length) {
    sh.getRange(head + 1, 1, rows.length, 5).setValues(rows);
    sh.getRange(head + 1, 4, rows.length, 2).setNumberFormat('#,##0.00');
    sh.getRange(head + 1, 1, rows.length, 1).setHorizontalAlignment('center');
    sh.getRange(head + 1, 3, rows.length, 1).setHorizontalAlignment('center');
  }
  const tr = head + rows.length + 1;
  sh.getRange(tr, 1, 1, 5).setBackground(SOFT).setFontWeight('bold');
  sh.getRange(tr, 2).setValue(L.total);
  sh.getRange(tr, 4).setValue(q.currency || '');
  sh.getRange(tr, 5).setValue(Number(q.total) || 0).setNumberFormat('#,##0.00');
  sh.getRange(head, 1, rows.length + 2, 5).setBorder(true, true, true, true, true, true, '#d8cfbd', SpreadsheetApp.BorderStyle.SOLID);

  const notes = (q.notes || []).filter(n => String(n || '').trim());
  if (notes.length) {
    sh.getRange(tr + 2, 1).setValue(L.notes).setFontWeight('bold').setFontColor(GOLD);
    notes.forEach((n, i) => sh.getRange(tr + 3 + i, 2, 1, 4).merge().setValue(n).setWrap(true));
  }
  trimSheet(sh, 6, tr + 3 + notes.length);
  sh.setColumnWidth(1, 45); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 70); sh.setColumnWidth(4, 110); sh.setColumnWidth(5, 120);
}

function deleteQuote(id) {
  const ss = SpreadsheetApp.getActive();
  const idx = indexSheet();
  const row = findRow(idx, id);
  if (row < 0) return;
  const tab = ss.getSheetByName(idx.getRange(row, 8).getValue());
  if (tab) ss.deleteSheet(tab);
  idx.deleteRow(row);
}

/* ---------- called from the app when it runs from this sheet ---------- */
function apiList() { return { ok: true, quotes: listQuotes() }; }
function apiSave(q) { return withLock(() => { saveQuote(q); return { ok: true }; }); }
function apiDelete(id) { return withLock(() => { deleteQuote(id); return { ok: true }; }); }
function withLock(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

/** Saves an exported PDF / image to a Drive folder next to the sheet, named after the client. */
function apiSaveFile(name, type, b64) {
  const folderName = 'MB System - ملفات العروض';
  const it = DriveApp.getFoldersByName(folderName);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
  const old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), type, name));
  return file.getUrl();
}


/** Keep each quote tab small (a new tab starts with 26,000 empty cells; Sheets allows 10 million per file). */
function trimSheet(sh, cols, rows) {
  const mr = sh.getMaxRows(), mc = sh.getMaxColumns();
  if (mr > rows) sh.deleteRows(rows + 1, mr - rows);
  else if (mr < rows) sh.insertRowsAfter(mr, rows - mr);
  if (mc > cols) sh.deleteColumns(cols + 1, mc - cols);
}
