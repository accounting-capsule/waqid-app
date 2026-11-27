/* Waqid sync. Paste into Apps Script of the Google Sheet. Deploy as Web app:
   Execute as: Me. Who has access: Anyone. Then run setup() once and set the key. */
var SNAP = '_snapshot', PREV = '_snapshot_prev', CHUNK = 40000;

function setup() { // run once: creates the key. Read it in Project settings > Script properties (never logged), paste it into the app once.
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SYNC_KEY')) p.setProperty('SYNC_KEY', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
}
function rotateKey() { // replaces the key; the old key stops working at once. Read the new one in Script properties
  PropertiesService.getScriptProperties().setProperty('SYNC_KEY', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return out({ ok: true, app: 'waqid-sync' }); }

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'bad_request' }); }
  var props = PropertiesService.getScriptProperties();
  var ownerKey = props.getProperty('SYNC_KEY'), clientKey = props.getProperty('CLIENT_KEY'), role = '';
  if (typeof req.key === 'string' && req.key) { if (ownerKey && req.key === ownerKey) role = 'owner'; else if (clientKey && req.key === clientKey) role = 'client'; }
  if (!role) return out({ ok: false, error: 'unauthorized' });
  if (role === 'client') return clientAppend_(req); // the client key can only add rows; it never reaches the code below
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rev = readRev_(ss);
    if (req.action === 'ping') return out({ ok: true, rev: rev });
    if (req.action === 'pull') return out({ ok: true, rev: rev, db: readSnap_(ss, SNAP) });
    if (req.action === 'inbox_pull') return out(inboxPull_(ss, 200));
    if (req.action === 'inbox_ack') return out(inboxAck_(ss, req.items));
    if (req.action === 'push') {
      if (+req.baseRev !== rev) return out({ ok: false, error: 'conflict', rev: rev });
      var db = req.db;
      if (!db || !Array.isArray(db.clients) || !Array.isArray(db.contracts) || !Array.isArray(db.bills) || !Array.isArray(db.journal)) return out({ ok: false, error: 'bad_data' });
      var cur = readSnap_(ss, SNAP);
      if (cur) writeSnap_(ss, PREV, cur, rev); // keep the previous version as a safety copy
      writeSnap_(ss, SNAP, db, rev + 1);
      writeTable_(ss, 'عملاء', ['الاسم', 'الهاتف', 'الجهة', 'العنوان', 'تاريخ الاتفاق', 'تاريخ الانتهاء', 'ملاحظات', 'ملاحظة النظام'], db.clients, ['name', 'phone', 'entity', 'address', 'agreeDate', 'endDate', 'notes', 'clientDeleted']);
      writeTable_(ss, 'عقود', null, db.contracts, null);
      writeTable_(ss, 'مستخلصات', null, db.bills, null);
      writeTable_(ss, 'حركات', null, db.journal, null);
      writeTable_(ss, 'خدمات', ['الخدمة'], (db.services || []).map(function (s) { return { s: s }; }), ['s']);
      return out({ ok: true, rev: rev + 1 });
    }
    return out({ ok: false, error: 'unknown_action' });
  } finally { lock.releaseLock(); }
}

function sheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function readRev_(ss) { var s = ss.getSheetByName(SNAP); return s ? (+s.getRange(1, 1).getValue() || 0) : 0; }
function writeSnap_(ss, name, db, rev) {
  var s = sheet_(ss, name); s.clear();
  var txt = JSON.stringify(db), rows = [[rev]];
  for (var i = 0; i < txt.length; i += CHUNK) rows.push([txt.substr(i, CHUNK)]);
  s.getRange(1, 1, rows.length, 1).setValues(rows);
  if (name === SNAP) s.hideSheet();
}
function readSnap_(ss, name) {
  var s = ss.getSheetByName(name); if (!s) return null;
  var n = s.getLastRow(); if (n < 2) return null;
  var v = s.getRange(2, 1, n - 1, 1).getValues(), t = '';
  for (var i = 0; i < v.length; i++) t += v[i][0];
  try { return JSON.parse(t); } catch (x) { return null; }
}
function writeTable_(ss, name, headers, rows, keys) {
  if (!keys) { // use the union of fields of the records
    var set = {}; rows.forEach(function (r) { Object.keys(r).forEach(function (k) { if (k.charAt(0) !== '_') set[k] = 1; }); });
    keys = Object.keys(set); headers = keys;
  }
  if (!headers.length) { headers = ['(فارغ)']; keys = []; rows = []; }
  var s = sheet_(ss, name); s.clear();
  var data = [headers].concat(rows.map(function (r) { return keys.map(function (k) { var v = r[k]; return v === undefined || v === null ? '' : v; }); }));
  s.getRange(1, 1, data.length, headers.length).setValues(data);
  s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0D1B2A').setFontColor('#C9A84C');
  s.setRightToLeft(true); s.setFrozenRows(1);
}


/* ---- Client inbox: a separate key that can only ADD rows. It cannot read, change or delete anything. ----
   Owner controls (run from the Apps Script editor, read the key in Project settings > Script properties):
   createClientKey()  makes the client key (does nothing if one exists)
   rotateClientKey()  replaces it; the old client key stops at once
   revokeClientKey()  deletes it; the client can no longer send. The owner key SYNC_KEY is never touched. */
function createClientKey() { var p = PropertiesService.getScriptProperties(); if (!p.getProperty('CLIENT_KEY')) rotateClientKey(); }
function rotateClientKey() { PropertiesService.getScriptProperties().setProperty('CLIENT_KEY', 'c' + Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '')); }
function revokeClientKey() { PropertiesService.getScriptProperties().deleteProperty('CLIENT_KEY'); }

var INBOX = {
  clients: { sheet: 'وارد_عملاء', head: ['الاسم', 'الهاتف', 'الجهة', 'العنوان', 'تاريخ الاتفاق', 'تاريخ الانتهاء', 'ملاحظات'], keys: ['name', 'phone', 'entity', 'address', 'agreeDate', 'endDate', 'notes'], num: [] },
  contracts: { sheet: 'وارد_عقود', head: ['اسم المشروع', 'العميل', 'نوع الخدمة', 'تاريخ العقد', 'الحالة', 'قيمة العقد', 'نسبة المحتجز', 'ملاحظات'], keys: ['name', 'client', 'service', 'date', 'status', 'value', 'retention', 'notes'], num: ['value', 'retention'] },
  bills: { sheet: 'وارد_مستخلصات', head: ['المشروع', 'البند', 'الرقم', 'التاريخ', 'مدة السداد', 'القيمة', 'ملاحظات'], keys: ['project', 'item', 'no', 'date', 'days', 'value', 'notes'], num: ['days', 'value'] },
  journal: { sheet: 'وارد_حركات', head: ['التاريخ', 'النوع', 'المشروع', 'البند', 'المبلغ', 'طريقة الدفع', 'رقم المستند', 'ملاحظات'], keys: ['date', 'type', 'project', 'item', 'amount', 'method', 'docNo', 'notes'], num: ['amount'] }
};
var LIMITS = { events: 50, perMinute: 20, rowsPerSheet: 20000, text: 500 };

function clientAppend_(req) {
  if (req.action !== 'append') return out({ ok: false, error: 'forbidden' });
  var ev = req.events;
  if (!Array.isArray(ev) || ev.length > LIMITS.events) return out({ ok: false, error: 'bad_data' });
  var cache = CacheService.getScriptCache(), n = +(cache.get('cl_rate') || 0);
  if (n >= LIMITS.perMinute) return out({ ok: false, error: 'rate' });
  cache.put('cl_rate', String(n + 1), 60);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), byKind = {}, accepted = 0, now = new Date();
    for (var i = 0; i < ev.length; i++) {
      var e = ev[i], k = e && INBOX[e.kind] ? e.kind : null;
      if (!k || (e.op !== 'upsert' && e.op !== 'delete') || !/^[A-Za-z0-9_-]{6,40}$/.test(String(e.evid)) || !/^[A-Za-z0-9_-]{4,40}$/.test(String(e.id))) continue;
      var cfg = INBOX[k], row = [now, e.evid, e.op === 'delete' ? 'حذف' : 'إضافة/تعديل', e.id], rec = e.rec && typeof e.rec === 'object' ? e.rec : {};
      for (var j = 0; j < cfg.keys.length; j++) row.push(e.op === 'delete' ? '' : cell_(rec[cfg.keys[j]], cfg.num.indexOf(cfg.keys[j]) >= 0));
      (byKind[k] = byKind[k] || []).push(row);
    }
    for (var kind in byKind) {
      var cf = INBOX[kind], s = inboxSheet_(ss, cf), last = s.getLastRow(), seen = {};
      if (last + byKind[kind].length > LIMITS.rowsPerSheet) return out({ ok: false, error: 'full' });
      if (last > 1) s.getRange(2, 2, last - 1, 1).getValues().forEach(function (v) { seen[v[0]] = 1; });
      var rows = byKind[kind].filter(function (r) { if (seen[r[1]]) return false; seen[r[1]] = 1; return true; });
      if (!rows.length) continue;
      var at = last + 1, w = rows[0].length;
      for (var c = 2; c <= w; c++) { // text columns get plain-text format, so nothing typed by the client can run as a formula
        var kn = c > 4 ? cf.keys[c - 5] : ''; if (cf.num.indexOf(kn) < 0) s.getRange(at, c, rows.length, 1).setNumberFormat('@');
      }
      s.getRange(at, 1, rows.length, 1).setValues(rows.map(function (r) { return [r[0]]; }));
      s.getRange(at, 2, rows.length, w - 1).setValues(rows.map(function (r) { return r.slice(1); }));
      accepted += rows.length;
    }
    return out({ ok: true, n: accepted });
  } finally { lock.releaseLock(); }
}
function cell_(v, isNum) {
  if (isNum) { var x = parseFloat(v); return isFinite(x) ? x : ''; }
  return v === undefined || v === null ? '' : String(v).slice(0, LIMITS.text);
}
function inboxSheet_(ss, cf) {
  var s = ss.getSheetByName(cf.sheet);
  if (!s) {
    s = ss.insertSheet(cf.sheet);
    var head = ['وقت الاستلام', 'رقم الحدث', 'العملية', 'معرف السجل'].concat(cf.head);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#0D1B2A').setFontColor('#C9A84C');
    s.setRightToLeft(true); s.setFrozenRows(1);
  }
  return s;
}


/* ---- Import (owner only): the app reads the events nobody has imported yet, merges them, then confirms. ----
   Rows stay in the inbox tabs as a log. The last column gets the import time once the app has confirmed. */
var MARK = 'تم الاستيراد';
function markCol_(cf) { return 5 + cf.keys.length; }
function inboxPull_(ss, limit) {
  var events = [], more = false;
  for (var kind in INBOX) {
    var cf = INBOX[kind], s = ss.getSheetByName(cf.sheet); if (!s) continue;
    var last = s.getLastRow(), mc = markCol_(cf); if (last < 2) continue;
    if (s.getRange(1, mc).getValue() !== MARK) s.getRange(1, mc).setValue(MARK).setFontWeight('bold').setBackground('#0D1B2A').setFontColor('#C9A84C');
    var marks = s.getRange(2, mc, last - 1, 1).getValues(), first = -1;
    for (var i = 0; i < marks.length; i++) if (marks[i][0] === '') { first = i; break; }
    if (first < 0) continue;
    var rows = s.getRange(2 + first, 1, last - 1 - first, mc - 1).getValues();
    for (var r = 0; r < rows.length; r++) {
      if (marks[first + r][0] !== '') continue;
      if (events.length >= limit) { more = true; break; }
      var row = rows[r], rec = {};
      if (row[2] !== 'حذف') for (var j = 0; j < cf.keys.length; j++) { var v = row[4 + j]; rec[cf.keys[j]] = (v instanceof Date) ? Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : v; }
      events.push({ kind: kind, evid: String(row[1]), op: row[2] === 'حذف' ? 'delete' : 'upsert', id: String(row[3]), rec: rec });
    }
  }
  return { ok: true, events: events, more: more };
}
function inboxAck_(ss, items) {
  if (!Array.isArray(items) || items.length > 500) return { ok: false, error: 'bad_data' };
  var by = {}, n = 0; items.forEach(function (it) { if (it && INBOX[it.kind] && typeof it.evid === 'string') (by[it.kind] = by[it.kind] || {})[it.evid] = 1; });
  for (var kind in by) {
    var cf = INBOX[kind], s = ss.getSheetByName(cf.sheet); if (!s) continue;
    var last = s.getLastRow(), mc = markCol_(cf); if (last < 2) continue;
    var ids = s.getRange(2, 2, last - 1, 1).getValues(), marks = s.getRange(2, mc, last - 1, 1).getValues(), now = new Date();
    for (var i = 0; i < ids.length; i++) if (by[kind][String(ids[i][0])] && marks[i][0] === '') { s.getRange(2 + i, mc).setValue(now); n++; }
  }
  return { ok: true, n: n };
}
