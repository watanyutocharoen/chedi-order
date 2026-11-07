/**
 * ระบบสั่งของงานเจดีย์ปากน้ำ 2569 — ฝั่งเซิร์ฟเวอร์ (Google Apps Script)
 * เก็บข้อมูลใน Google Sheet ที่สคริปต์นี้ผูกอยู่ (แท็บ people, shops, items, orders)
 * ใช้คู่กับไฟล์ Index.html
 */
var COLS = ['people', 'shops', 'items', 'orders'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('สั่งของงานเจดีย์ปากน้ำ 2569')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() {
  var s = SpreadsheetApp.getActiveSpreadsheet();
  if (s) return s;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SSID');
  if (id) return SpreadsheetApp.openById(id);
  s = SpreadsheetApp.create('ข้อมูลสั่งของงานเจดีย์ปากน้ำ');
  props.setProperty('SSID', s.getId());
  return s;
}

function sheet_(name) {
  if (COLS.indexOf(name) < 0) throw new Error('ไม่รู้จักตาราง ' + name);
  var s = ss_(), sh = s.getSheetByName(name);
  if (!sh) {
    sh = s.insertSheet(name);
    sh.appendRow(['id', 'json']);
    sh.setColumnWidth(2, 500);
  }
  return sh;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/** อ่านข้อมูลทั้งหมดในครั้งเดียว */
function getAll() {
  var out = {};
  COLS.forEach(function (n) {
    var sh = sheet_(n), last = sh.getLastRow();
    var rows = last > 1 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
    out[n] = rows.map(function (r) {
      var o = {};
      try { o = JSON.parse(r[1]); } catch (e) {}
      o.id = String(r[0]);
      return o;
    });
  });
  return out;
}

function addDoc(coll, json) {
  return withLock_(function () {
    var id = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    sheet_(coll).appendRow([id, json]);
    return id;
  });
}

function updateDoc(coll, id, json) {
  return withLock_(function () {
    var sh = sheet_(coll), row = findRow_(sh, id);
    if (row < 0) throw new Error('ไม่พบรายการนี้แล้ว (อาจถูกลบไปก่อน)');
    var cur = {};
    try { cur = JSON.parse(sh.getRange(row, 2).getValue() || '{}'); } catch (e) {}
    var upd = JSON.parse(json);
    Object.keys(upd).forEach(function (k) {
      if (upd[k] && typeof upd[k] === 'object' && upd[k].__delete__) delete cur[k];
      else cur[k] = upd[k];
    });
    sh.getRange(row, 2).setValue(JSON.stringify(cur));
    return true;
  });
}

function deleteDoc(coll, id) {
  return withLock_(function () {
    var sh = sheet_(coll), row = findRow_(sh, id);
    if (row > 0) sh.deleteRow(row);
    return true;
  });
}

/** ใส่ข้อมูลตั้งต้น (รายชื่อ ร้าน เมนู) ครั้งแรกเท่านั้น ถ้าตารางว่าง */
function seedAll(json) {
  return withLock_(function () {
    if (sheet_('shops').getLastRow() > 1 || sheet_('people').getLastRow() > 1) return false;
    var d = JSON.parse(json);
    COLS.forEach(function (n) {
      var sh = sheet_(n);
      var rows = (d[n] || []).map(function (o) {
        var id = o.id, c = {};
        Object.keys(o).forEach(function (k) { if (k !== 'id') c[k] = o[k]; });
        return [id, JSON.stringify(c)];
      });
      if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
    });
    return true;
  });
}
