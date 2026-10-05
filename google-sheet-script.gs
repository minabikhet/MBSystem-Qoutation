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
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('MB System Quotations')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  } catch (err) {
    return out({ ok: true });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'save') { saveQuote(body.quote); safeProfit(); return out({ ok: true }); }
    if (body.action === 'delete') { deleteQuote(body.id); safeProfit(); return out({ ok: true }); }
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
  const base = (q.type === 'invoice' ? 'فاتورة - ' : q.type === 'purchase' ? 'شراء - ' : '') + cleanName(q.client) + ' - ' + dmy(q.date).replace(/\//g, '-');
  let tabName = oldTab && oldTab.indexOf(base) === 0 ? oldTab : uniqueName(base, oldTab);

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

const LOGO_URL = 'https://minabikhet.github.io/MBSystem-Qoutation/logo.png';
const PHONES = ['01017886248', '01222455495'];
const SITE = 'https://minabikhet.github.io/MBSystem-Qoutation/';
const SIG_URL = SITE + 'signature.png';
const ICON_PW = SITE + 'icon-phone-wa.png';
const ICON_FB = SITE + 'icon-fb.png';
const FB_URL = 'https://www.facebook.com/share/1CAqvtfLLw/';

/** Draws the quote tab to look like the printed quotation in the app. */
function writeQuoteTab(sh, q) {
  const en = q.lang === 'en';
  const L = en
    ? { title: 'QUOTATION', client: 'Client:', date: 'Date:', greet: 'We are pleased to submit the following quotation:', no: '#', desc: 'Description', qty: 'Qty', price: 'Unit Price', total: 'Total', notes: 'Notes:', contact: 'Contact Us:', regards: 'Best regards' }
    : { title: 'عـرض سعـر', client: 'اسم العميل /', date: 'التاريخ:', greet: 'تحية طيبة وبعد، نتشرف بتقديم عرض السعر التالي:', no: 'م', desc: 'البيان والمواصفات', qty: 'الكمية', price: 'السعر', total: 'الإجمالي', notes: 'ملاحظات:', contact: 'Contact US :', regards: 'مع خالص التحية والشكر' };
  const items = (q.items || []).filter(it => String(it.d || '').trim());
  if (q.type === 'purchase') {
    L.title = en ? 'PURCHASE INVOICE' : 'فاتورة شراء';
    L.greet = en ? 'Purchase cost of the supplies for this client:' : 'بيان بأسعار شراء المستلزمات الخاصة بالعميل:';
  }
  if (q.type === 'invoice') {
    L.title = en ? 'INVOICE' : 'فـاتـورة';
    L.greet = en ? 'Please find below the invoice for the following works and supplies:' : 'تحية طيبة وبعد، مرفق لسيادتكم الفاتورة الخاصة بالأعمال والمشتريات التالية:';
  }
  const notes = q.showNotes === false ? [] : (q.notes || []).filter(n => String(n || '').trim());
  const LINE = '#E6E1D6';
  const S = SpreadsheetApp.BorderStyle;

  // reset the tab
  const need = 44 + items.length + notes.length;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < 7) sh.insertColumnsAfter(sh.getMaxColumns(), 7 - sh.getMaxColumns());
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart(); sh.clear(); all.setBorder(false, false, false, false, false, false);
  sh.setRightToLeft(!en);
  sh.setHiddenGridlines(true);
  all.setFontFamily('Cairo').setFontSize(10).setFontColor(INK_()).setVerticalAlignment('middle');
  const end = en ? 'right' : 'left', start = en ? 'left' : 'right';

  // columns: margin | # | description | qty | price | total | margin
  [14, 42, 330, 70, 105, 125, 14].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // header: title block + logo
  sh.setRowHeight(1, 16);
  [2, 3, 4, 5].forEach(r => sh.setRowHeight(r, 26));
  M(sh.getRange('B2:D3')).setValue(L.title).setFontSize(22).setFontWeight('bold').setFontColor(NAVY).setHorizontalAlignment('center');
  M(sh.getRange('B4:D4')).setValue('MB SYSTEM').setFontFamily('Arial').setFontSize(13).setFontWeight('bold').setFontColor(GOLD).setHorizontalAlignment('center');
  M(sh.getRange('B5:D5')).setValue('SECURITY SOLUTIONS').setFontFamily('Arial').setFontSize(8).setFontColor('#5d6773').setHorizontalAlignment('center').setVerticalAlignment('top');
  M(sh.getRange('E2:F5')).setFormula('=IMAGE("' + LOGO_URL + '",1)');
  sh.setRowHeight(6, 8);
  sh.getRange('B6:F6').setBorder(null, null, true, null, null, null, GOLD, S.SOLID_MEDIUM);

  // client + date
  sh.setRowHeight(7, 30);
  const cname = String(q.client || '');
  let rt = SpreadsheetApp.newRichTextValue().setText(L.client + ' ' + cname);
  if (cname) rt = rt.setTextStyle(L.client.length + 1, L.client.length + 1 + cname.length,
    SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(NAVY).build());
  M(sh.getRange('B7:C7')).setRichTextValue(rt.build()).setFontSize(11).setHorizontalAlignment(start);
  M(sh.getRange('D7:F7')).setValue((q.no ? q.no + '   ·   ' : '') + L.date + ' ' + dmy(q.date)).setFontSize(11).setFontWeight('bold').setHorizontalAlignment(end);
  M(sh.getRange('B8:F8')).setValue(L.greet).setFontColor('#3a4552').setHorizontalAlignment(start);

  // table
  const head = 9;
  sh.setRowHeight(head, 30);
  sh.getRange(head, 2, 1, 5).setValues([[L.no, L.desc, L.qty, L.price, L.total]])
    .setFontWeight('bold').setBackground(SOFT).setFontColor(NAVY).setHorizontalAlignment('center')
    .setBorder(null, null, true, null, null, null, GOLD, S.SOLID_MEDIUM);
  const rows = items.map((it, i) => [i + 1, it.d, num(it.q), num(it.p), Math.round(num(it.q) * num(it.p) * 100) / 100]);
  const pad = Math.max(0, 12 - rows.length - notes.length);
  const bodyN = rows.length + pad;
  if (rows.length) sh.getRange(head + 1, 2, rows.length, 5).setValues(rows);
  if (bodyN) {
    const body = sh.getRange(head + 1, 2, bodyN, 5);
    body.setHorizontalAlignment('center').setBorder(null, null, null, null, null, true, LINE, S.SOLID);
    sh.getRange(head + bodyN, 2, 1, 5).setBorder(null, null, true, null, null, null, LINE, S.SOLID);
    sh.getRange(head + 1, 3, bodyN, 1).setHorizontalAlignment(start);
    sh.getRange(head + 1, 4, bodyN, 1).setNumberFormat('#,##0.##');
    sh.getRange(head + 1, 5, bodyN, 2).setNumberFormat('#,##0.00');
    for (let r = head + 1; r <= head + bodyN; r++) sh.setRowHeight(r, 26);
  }

  // total
  let tr = head + bodyN + 1;
  if (q.vat === true) {
    const sub = Number(q.sub) || 0, vat = Number(q.vatAmount) || 0;
    [[en ? 'Subtotal' : 'الإجمالي قبل الضريبة', sub], [en ? 'VAT 14%' : 'ضريبة القيمة المضافة 14%', vat]].forEach(([lab, val]) => {
      M(sh.getRange(tr, 2, 1, 4)).setValue(lab).setHorizontalAlignment(start);
      sh.getRange(tr, 6).setValue(val).setNumberFormat('#,##0.00').setFontWeight('bold').setHorizontalAlignment('center');
      sh.getRange(tr, 2, 1, 5).setBorder(null, null, true, null, null, null, '#E6E1D6', S.SOLID);
      sh.setRowHeight(tr, 26); tr++;
    });
    L.total = en ? 'Total incl. VAT' : 'الإجمالي شامل الضريبة';
  }
  sh.setRowHeight(tr, 38);
  const tot = sh.getRange(tr, 2, 1, 5).setBackground(SOFT).setBorder(true, null, null, null, null, null, GOLD, S.SOLID_MEDIUM);
  M(sh.getRange(tr, 2, 1, 3)).setValue(L.total).setFontSize(13).setFontWeight('bold').setFontColor(NAVY).setHorizontalAlignment(start);
  sh.getRange(tr, 5).setValue(q.currency || '').setFontSize(9).setFontColor('#5d6773').setHorizontalAlignment(end);
  sh.getRange(tr, 6).setValue(Number(q.total) || 0).setNumberFormat('#,##0.00').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

  // notes
  let r = tr + 2;
  if (notes.length) {
    M(sh.getRange(r, 2, 1, 5)).setValue(L.notes).setFontWeight('bold').setFontColor(GOLD).setHorizontalAlignment(start);
    r++;
    notes.forEach(n => {
      M(sh.getRange(r, 2, 1, 5)).setValue(n).setWrap(true).setFontColor('#2c3540').setHorizontalAlignment(start);
      sh.setRowHeight(r, 30); r++;
    });
  }

  // footer
  r += 1;
  const ft = r;
  sh.getRange(ft, 2, 1, 5).setBorder(true, null, null, null, null, null, GOLD, S.SOLID_MEDIUM);
  M(sh.getRange(ft, 2, 1, 2)).setValue(L.contact).setFontFamily('Arial').setFontWeight('bold').setFontColor(NAVY).setHorizontalAlignment(start);
  // phone + WhatsApp icons in column B, the number beside them in column C
  [PHONES[0], PHONES[1]].forEach((p, i) => {
    sh.getRange(ft + 1 + i, 2).setFormula('=IMAGE("' + ICON_PW + '",1)');
    sh.getRange(ft + 1 + i, 3).setValue("'" + p).setFontFamily('Arial').setFontWeight('bold').setHorizontalAlignment(start);
  });
  sh.getRange(ft + 3, 2).setFormula('=IMAGE("' + ICON_FB + '",1)');
  sh.getRange(ft + 3, 3).setFormula('=HYPERLINK("' + FB_URL + '","' + String(q.fb || 'MB Systems').replace(/"/g, '""') + '")')
    .setFontFamily('Arial').setFontWeight('bold').setFontColor('#1f5fbf').setHorizontalAlignment(start);
  [ft + 1, ft + 2, ft + 3].forEach(rr => sh.setRowHeight(rr, 22));
  M(sh.getRange(ft, 4, 1, 3)).setValue(L.regards).setFontColor('#5d6773').setHorizontalAlignment(end);
  M(sh.getRange(ft + 1, 4, 1, 3)).setValue('MB Systems').setFontFamily('Arial').setFontWeight('bold').setFontColor(NAVY).setHorizontalAlignment(end);
  if (q.sign === true) {
    // signature replaces "Security Solutions"
    M(sh.getRange(ft + 2, 5, 2, 2)).setFormula('=IMAGE("' + SIG_URL + '",1)');
  } else {
    M(sh.getRange(ft + 2, 4, 1, 3)).setValue('Security Solutions').setFontFamily('Arial').setFontSize(9).setFontColor('#5d6773').setHorizontalAlignment(end);
  }
  const last = ft + 4;
  sh.setRowHeight(last, 16);

  // gold frame around the page
  sh.getRange(2, 2, last - 2, 5).setBorder(true, true, true, true, null, null, GOLD, S.SOLID);
  trimSheet(sh, 7, last + 1);
}
function INK_() { return '#1b2430'; }
/** Merge a block and hand back its top-left cell, where the value and formatting belong. */
function M(range) { range.merge(); return range.getCell(1, 1); }

/** Run once from the editor (choose rebuildAllTabs, then Run) to redraw old quote tabs in the new design. */
function rebuildAllTabs() {
  const ss = SpreadsheetApp.getActive();
  const idx = indexSheet();
  const n = idx.getLastRow();
  if (n < 2) return;
  const data = idx.getRange(2, 1, n - 1, 9).getValues();
  data.forEach(row => {
    let q; try { q = JSON.parse(row[8]); } catch (e) { return; }
    const sh = ss.getSheetByName(row[7]);
    if (sh) writeQuoteTab(sh, q);
  });
  safeProfit();
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

/* ---------- profit tab: sales invoices minus their purchase invoices, grouped by month ---------- */
const PROFIT = 'الأرباح';
function safeProfit() { try { buildProfitTab(); } catch (e) { console.error(e); } }
function netOf(x) { return Number(x.sub != null ? x.sub : x.total) || 0; }
function normName(n) { return String(n || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function buildProfitTab() {
  const ss = SpreadsheetApp.getActive();
  const all = listQuotes();
  const sales = all.filter(x => x.type === 'invoice');
  const buys = all.filter(x => x.type === 'purchase');
  const deals = {};
  sales.forEach(s => deals[s.id] = { s: s, cost: 0, n: 0 });
  const orphans = [];
  buys.forEach(p => {
    let id = p.linkTo && deals[p.linkTo] ? p.linkTo : '';
    if (!id) {
      const c = sales.filter(s => normName(s.client) === normName(p.client));
      if (c.length) {
        const pd = Date.parse(p.date) || 0;
        c.sort((a, b) => Math.abs((Date.parse(a.date) || 0) - pd) - Math.abs((Date.parse(b.date) || 0) - pd));
        id = c[0].id;
      }
    }
    if (id) { deals[id].cost += netOf(p); deals[id].n++; } else orphans.push(p);
  });
  const months = {};
  const M = k => months[k] || (months[k] = { sales: 0, cost: 0, rows: [] });
  Object.keys(deals).forEach(id => {
    const d = deals[id], m = M(String(d.s.date || '').slice(0, 7)), sale = netOf(d.s);
    if (!d.n) { m.rows.push([dmy(d.s.date), d.s.client, sale, '', '', 'مستنية فاتورة الشراء (مش محسوبة)']); return; }
    m.sales += sale; m.cost += d.cost;
    m.rows.push([dmy(d.s.date), d.s.client, sale, d.cost, sale - d.cost, d.s.no || '']);
  });
  orphans.forEach(p => {
    const m = M(String(p.date || '').slice(0, 7)), c = netOf(p);
    m.rows.push([dmy(p.date), p.client, '', c, '', 'مستنية فاتورة البيع (مش محسوبة)']);
  });

  let sh = ss.getSheetByName(PROFIT);
  if (!sh) sh = ss.insertSheet(PROFIT, 1);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.setRightToLeft(true);
  sh.setHiddenGridlines(true);
  [110, 200, 120, 120, 120, 190].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const keys = Object.keys(months).sort().reverse();
  let S = 0, C = 0; keys.forEach(k => { S += months[k].sales; C += months[k].cost; });
  const out = [];
  const fmtRows = [];
  out.push(['إجمالي المكسب', '', S, C, S - C, '']); fmtRows.push(['total', out.length]);
  out.push(['', '', 'المبيعات', 'المشتريات', 'المكسب', '']); fmtRows.push(['sub', out.length]);
  out.push(['', '', '', '', '', '']);
  keys.forEach(k => {
    const m = months[k];
    out.push([k, '', m.sales, m.cost, m.sales - m.cost, m.sales ? Math.round((m.sales - m.cost) / m.sales * 100) + '%' : '']); fmtRows.push(['month', out.length]);
    out.push(['التاريخ', 'العميل', 'البيع', 'الشراء', 'المكسب', 'ملاحظة']); fmtRows.push(['head', out.length]);
    m.rows.forEach(r => out.push(r));
    out.push(['', '', '', '', '', '']);
  });
  const need = out.length + 2;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(1, 1, out.length, 6).setValues(out).setFontFamily('Cairo').setVerticalAlignment('middle');
  sh.getRange(1, 3, out.length, 3).setNumberFormat('#,##0.00').setHorizontalAlignment('center');
  fmtRows.forEach(([kind, r]) => {
    const rg = sh.getRange(r, 1, 1, 6);
    if (kind === 'total') { rg.setBackground(NAVY).setFontColor('#ffffff').setFontWeight('bold').setFontSize(13); sh.setRowHeight(r, 34); }
    if (kind === 'sub') rg.setFontColor('#5d6773').setFontSize(9).setHorizontalAlignment('center');
    if (kind === 'month') { rg.setBackground(SOFT).setFontWeight('bold').setFontColor(NAVY).setBorder(true, null, true, null, null, null, GOLD, SpreadsheetApp.BorderStyle.SOLID_MEDIUM); sh.setRowHeight(r, 30); }
    if (kind === 'head') rg.setFontWeight('bold').setFontColor('#5d6773');
  });
  sh.setFrozenRows(2);
}
