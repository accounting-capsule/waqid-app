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
  var key = PropertiesService.getScriptProperties().getProperty('SYNC_KEY');
  if (!key || req.key !== key) return out({ ok: false, error: 'unauthorized' });
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rev = readRev_(ss);
    if (req.action === 'ping') return out({ ok: true, rev: rev });
    if (req.action === 'pull') return out({ ok: true, rev: rev, db: readSnap_(ss, SNAP) });
    if (req.action === 'push') {
      if (+req.baseRev !== rev) return out({ ok: false, error: 'conflict', rev: rev });
      var db = req.db;
      if (!db || !Array.isArray(db.clients) || !Array.isArray(db.contracts) || !Array.isArray(db.bills) || !Array.isArray(db.journal)) return out({ ok: false, error: 'bad_data' });
      var cur = readSnap_(ss, SNAP);
      if (cur) writeSnap_(ss, PREV, cur, rev); // keep the previous version as a safety copy
      writeSnap_(ss, SNAP, db, rev + 1);
      writeTable_(ss, 'عملاء', ['الاسم', 'الهاتف', 'الجهة', 'العنوان', 'تاريخ الاتفاق', 'تاريخ الانتهاء', 'ملاحظات'], db.clients, ['name', 'phone', 'entity', 'address', 'agreeDate', 'endDate', 'notes']);
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
    var set = {}; rows.forEach(function (r) { Object.keys(r).forEach(function (k) { set[k] = 1; }); });
    keys = Object.keys(set); headers = keys;
  }
  if (!headers.length) { headers = ['(فارغ)']; keys = []; rows = []; }
  var s = sheet_(ss, name); s.clear();
  var data = [headers].concat(rows.map(function (r) { return keys.map(function (k) { var v = r[k]; return v === undefined || v === null ? '' : v; }); }));
  s.getRange(1, 1, data.length, headers.length).setValues(data);
  s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0D1B2A').setFontColor('#C9A84C');
  s.setRightToLeft(true); s.setFrozenRows(1);
}
