/**
 * BRS Quotation Center — ตัวกลางระหว่างหน้าเว็บกับ Google Sheets / Google Drive
 *
 * วิธีติดตั้ง (ดูคู่มือ README ประกอบ):
 *   1) วางโค้ดนี้ทั้งหมดใน Extensions > Apps Script ของ Google Sheet ที่สร้างไว้
 *   2) เลือกฟังก์ชัน importOldData แล้วกด Run หนึ่งครั้ง (สร้างชีต โฟลเดอร์ และย้ายข้อมูลเดิม; อนุญาตสิทธิ์ตามที่ Google ถาม)
 *   3) Deploy > New deployment > Web app  (Execute as: Me, Who has access: Anyone)
 *
 * ข้อมูลจริงอยู่ในคอลัมน์สุดท้าย "data" ของชีต Quotes  คอลัมน์อื่นมีไว้ให้คนอ่านเท่านั้น
 * ห้ามแก้ข้อมูลในชีตด้วยมือ ให้แก้ผ่านหน้าเว็บ
 */

var SHEET_QUOTES = 'Quotes';
var SHEET_NUMBERS = 'Numbers';
var SHEET_SETTINGS = 'Settings';
var FOLDER_NAME = 'BRS Quotation Center - เอกสารที่เซ็นแล้ว';
var DEFAULT_PREFIX = 'Q-SCO-';

var Q_HEAD = ['id', 'เลขที่', 'วันที่', 'วันหมดอายุ', 'บริษัท', 'ผู้ติดต่อ', 'ยอดรวม (บาท)', 'VAT', 'ผู้จัดทำ',
  'อีเมลผู้จัดทำ', 'เบอร์ผู้จัดทำ', 'ยกเลิก', 'ไฟล์ที่เซ็นแล้ว', 'สร้างเมื่อ', 'แก้ไขล่าสุด', 'data (ห้ามแก้)'];
var N_HEAD = ['เลขที่', 'ปี', 'ลำดับ', 'id ใบเสนอราคา', 'ออกเลขเมื่อ'];
var S_HEAD = ['key', 'value'];

/* ------------------------------------------------------------------ setup */
function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET_QUOTES, Q_HEAD);
  ensureSheet_(ss, SHEET_NUMBERS, N_HEAD);
  ensureSheet_(ss, SHEET_SETTINGS, S_HEAD);
  var q = ss.getSheetByName(SHEET_QUOTES);
  q.getRange(1, 1, q.getMaxRows(), Q_HEAD.length).setNumberFormat('@');
  q.getRange(2, 7, Math.max(1, q.getMaxRows() - 1), 1).setNumberFormat('#,##0.00');
  ss.getSheetByName(SHEET_NUMBERS).getRange('A:E').setNumberFormat('@');
  ss.getSheetByName(SHEET_SETTINGS).getRange('A:B').setNumberFormat('@');
  var folder = rootFolder_();
  Logger.log('พร้อมใช้งาน  โฟลเดอร์เก็บไฟล์: ' + folder.getUrl());
}

function ensureSheet_(ss, name, head) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#c4bd97');
  sh.setFrozenRows(1);
  return sh;
}

function rootFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function quoteFolder_(no) {
  var root = rootFolder_();
  var name = String(no || 'ไม่มีเลขที่');
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

/* -------------------------------------------------------------- web entry */
function doGet() {
  return json_({ ok: true, app: 'BRS Quotation Center', message: 'ระบบพร้อมใช้งาน' });
}

function doPost(e) {
  try {
    var p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (p.action) {
      case 'bootstrap': return json_(bootstrap_());
      case 'saveQuote': return json_(saveQuote_(p));
      case 'cancelQuote': return json_(cancelQuote_(p));
      case 'uploadFile': return json_(uploadFile_(p));
      case 'deleteFile': return json_(deleteFile_(p));
      case 'saveSettings': return json_(saveSettings_(p));
      default: throw fail_('bad_action', 'ไม่รู้จักคำสั่ง ' + p.action);
    }
  } catch (err) {
    return json_({ ok: false, code: err.code || 'server', message: err.message || String(err), no: err.no });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fail_(code, message, extra) { var e = new Error(message); e.code = code; if (extra) for (var k in extra) e[k] = extra[k]; return e; }
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw fail_('busy', 'มีคนกำลังบันทึกอยู่ ลองอีกครั้ง');
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ------------------------------------------------------------------ reads */
function sheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw fail_('not_setup', 'ยังไม่ได้รันฟังก์ชัน setup ใน Apps Script');
  return sh;
}

function readQuotes_() {
  var sh = sheet_(SHEET_QUOTES), last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, Q_HEAD.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    var id = String(rows[i][0] || ''); if (!id) continue;
    var data; try { data = JSON.parse(rows[i][Q_HEAD.length - 1] || '{}'); } catch (e) { continue; }
    out.push({ id: id, row: i + 2, data: data });
  }
  return out;
}

function readNumbers_() {
  var sh = sheet_(SHEET_NUMBERS), last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, N_HEAD.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({ no: String(rows[i][0]), year: Number(rows[i][1]) || null, seq: Number(rows[i][2]) || null, quoteId: String(rows[i][3] || ''), at: String(rows[i][4] || '') });
  }
  return out;
}

function readSettings_() {
  var sh = sheet_(SHEET_SETTINGS), last = sh.getLastRow(), out = {};
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    try { out[rows[i][0]] = JSON.parse(rows[i][1]); } catch (e) { out[rows[i][0]] = rows[i][1]; }
  }
  return out;
}

function bootstrap_() {
  return {
    ok: true,
    quotes: readQuotes_().map(function (q) { return { id: q.id, data: q.data }; }),
    numbers: readNumbers_(),
    settings: readSettings_()
  };
}

/* ----------------------------------------------------------------- writes */
function rowFor_(id, q) {
  var files = (q.signedDocs || []).map(function (d) { return d.name + '  ' + (d.url || ''); }).join('\n');
  return [id, q.no || '', q.date || '', q.expireDate || '', q.company || '', q.contact || '', Number(q.total) || 0,
    q.vat ? 'มี' : '', q.preparer || '', q.prepEmail || '', q.prepTel || '', q.cancelled ? 'ยกเลิก' : '', files,
    q.createdAt || '', q.updatedAt || '', JSON.stringify(q)];
}

function writeQuote_(id, q, row) {
  var sh = sheet_(SHEET_QUOTES);
  var values = [rowFor_(id, q)];
  if (row) sh.getRange(row, 1, 1, Q_HEAD.length).setValues(values);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, Q_HEAD.length).setValues(values);
}

function parseNo_(no) {
  var m = String(no || '').match(/(\d{4})(\d{3,})\s*$/);
  return m ? { year: Number(m[1]), seq: Number(m[2]) } : { year: null, seq: null };
}
function pad3_(n) { var s = String(n); while (s.length < 3) s = '0' + s; return s; }

/* Numbering: prefix + year of the document date + running number (3+ digits).
   Every save re-reads the latest number of that year; every issued number is kept in "Numbers" and never reused. */
function saveQuote_(p) {
  return withLock_(function () {
    var q = p.quote || {};
    var quotes = readQuotes_(), numbers = readNumbers_(), settings = readSettings_();
    var prefix = settings.prefix != null && settings.prefix !== '' ? String(settings.prefix) : DEFAULT_PREFIX;
    var id = p.id || Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    var existing = null;
    for (var i = 0; i < quotes.length; i++) if (quotes[i].id === id) existing = quotes[i];
    if (existing && existing.data.cancelled) throw fail_('cancelled', 'ใบนี้ถูกยกเลิกแล้ว แก้ไขไม่ได้');

    var year = Number(String(q.date || '').slice(0, 4));
    if (!year) throw fail_('bad_date', 'กรอกวันที่ของเอกสาร');

    function taken(no) {
      no = String(no).trim();
      for (var a = 0; a < numbers.length; a++) if (numbers[a].no === no && numbers[a].quoteId !== id) return true;
      for (var b = 0; b < quotes.length; b++) if (quotes[b].id !== id && String(quotes[b].data.no || '').trim() === no) return true;
      return false;
    }
    var typed = String(q.no || '').trim();
    if (p.autoNo || !typed) {
      var max = 0;
      numbers.forEach(function (n) { if (n.year === year && n.seq > max) max = n.seq; });
      quotes.forEach(function (x) { if (Number(x.data.year) === year && Number(x.data.seq) > max) max = Number(x.data.seq); });
      var seq = max + 1;
      while (taken(prefix + year + pad3_(seq))) seq++;
      q.no = prefix + year + pad3_(seq); q.year = year; q.seq = seq;
    } else {
      var savedNo = existing ? String(existing.data.no || '').trim() : '';
      if (typed !== savedNo && taken(typed)) throw fail_('dup', 'เลขที่ ' + typed + ' เคยออกไปแล้ว', { no: typed });
      var pn = parseNo_(typed);
      q.no = typed; q.year = pn.year; q.seq = pn.seq;
    }

    var now = new Date().toISOString();
    q.signedDocs = existing ? (existing.data.signedDocs || []) : [];
    q.cancelled = false;
    q.createdAt = existing ? (existing.data.createdAt || now) : now;
    q.updatedAt = now;
    delete q.autoNo;
    writeQuote_(id, q, existing && existing.row);

    var number = null;
    var known = numbers.some(function (n) { return n.no === q.no; });
    if (!known) {
      number = { no: q.no, year: q.year, seq: q.seq, quoteId: id, at: now };
      var ns = sheet_(SHEET_NUMBERS);
      ns.getRange(ns.getLastRow() + 1, 1, 1, N_HEAD.length).setValues([[number.no, number.year || '', number.seq || '', id, now]]);
    }
    return { ok: true, id: id, quote: q, number: number };
  });
}

function findQuote_(id) {
  var quotes = readQuotes_();
  for (var i = 0; i < quotes.length; i++) if (quotes[i].id === id) return quotes[i];
  throw fail_('not_found', 'ไม่พบใบเสนอราคานี้');
}

function cancelQuote_(p) {
  return withLock_(function () {
    var x = findQuote_(p.id), now = new Date().toISOString();
    x.data.cancelled = true; x.data.cancelledAt = now; x.data.updatedAt = now;
    writeQuote_(x.id, x.data, x.row);
    return { ok: true, id: x.id, quote: x.data };
  });
}

function uploadFile_(p) {
  if (!p.b64 || !p.name) throw fail_('bad_file', 'ไม่มีไฟล์');
  var allowed = { 'application/pdf': 1, 'image/jpeg': 1, 'image/png': 1, 'image/webp': 1 };
  if (!allowed[p.type]) throw fail_('unsupported_type', 'รองรับเฉพาะ PDF, JPG, PNG');
  var head = findQuote_(p.id);
  if (head.data.cancelled) throw fail_('cancelled', 'ใบนี้ถูกยกเลิกแล้ว');
  var bytes = Utilities.base64Decode(p.b64);
  if (bytes.length > 20 * 1024 * 1024) throw fail_('too_large', 'ไฟล์ใหญ่เกิน 20 MB');
  var file = quoteFolder_(head.data.no).createFile(Utilities.newBlob(bytes, p.type, p.name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var entry = { id: file.getId(), name: p.name, type: p.type, size: bytes.length, at: new Date().toISOString(), url: file.getUrl() };
  return withLock_(function () {
    var x = findQuote_(p.id);
    x.data.signedDocs = (x.data.signedDocs || []).concat([entry]);
    writeQuote_(x.id, x.data, x.row);
    return { ok: true, id: x.id, quote: x.data, file: entry };
  });
}

function deleteFile_(p) {
  return withLock_(function () {
    var x = findQuote_(p.id);
    var before = (x.data.signedDocs || []).length;
    x.data.signedDocs = (x.data.signedDocs || []).filter(function (d) { return d.id !== p.fileId; });
    if (x.data.signedDocs.length !== before) {
      try { DriveApp.getFileById(p.fileId).setTrashed(true); } catch (e) {}
      writeQuote_(x.id, x.data, x.row);
    }
    return { ok: true, id: x.id, quote: x.data };
  });
}

function saveSettings_(p) {
  return withLock_(function () {
    var s = p.settings || {}, sh = sheet_(SHEET_SETTINGS);
    var keys = Object.keys(s).sort();
    var last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, 2).clearContent();
    if (keys.length) sh.getRange(2, 1, keys.length, 2).setValues(keys.map(function (k) { return [k, JSON.stringify(s[k])]; }));
    return { ok: true, settings: s };
  });
}

/* ------------------------------------------- one-time import of old data */
var OLD_DATA = {"quotes": [{"id": "a0jlqasr7l81e53hzbon", "data": {"addr1": "999 ถนนสามเสน แขวงถนนนครไชยศรี", "addr2": "เขตดุสิต กรุงเทพมหานคร 10300", "approver": "คุณปกรณ์ รัตนสุวรรณ", "cancelled": true, "cancelledAt": "2026-11-02T00:19:32.375Z", "company": "บริษัท บีเอสจี พลัส จำกัด", "contact": "ศรัณย์ เจริญสุข", "createdAt": "2026-10-31T08:42:08.856Z", "credit": "30 วัน", "date": "2026-10-31", "expireDate": "2026-12-31", "fax": "-", "headerNote": "", "items": [{"desc": "ค่าบริการขนกาก", "price": 70000, "qty": 1}, {"desc": "ค่าบริการติดต่อภายใน", "price": 3000, "qty": 1}], "no": "Q-SCO-2026013", "notes": ["ใช้ภายในเท่านั้น", "คิดเงินภายใน 30 วัน", "เซ็นเอกสารภายใน 30 วัน", "ติดต่อค่าใช้จ่ายภายใน 30 วัน"], "prepEmail": "napong_a@brs-group.com", "prepTel": "0882989941", "preparer": "ณพง เอกภาณุศิริ", "priceUnit": "บาท/เดือน", "seq": 13, "subtotal": 73000, "taxId": "1100200587223", "tel": "0870076004", "total": 73000, "updatedAt": "2026-11-02T00:19:32.375Z", "vat": false, "vatAmount": 0, "year": 2026}}, {"id": "seed-Q-SCO-2026012", "data": {"addr1": "999 ถนนสามเสน แขวงถนนนครไชยศรี", "addr2": "เขตดุสิต กรุงเทพมหานคร 10300", "approver": "คุณปกรณ์ รัตนสุวรรณ", "company": "บริษัท บีเอสจี พลัส จำกัด", "contact": "ศรัณย์ เจริญสุข", "createdAt": "2026-06-04T09:00:00.000Z", "credit": "30 วัน", "date": "2026-06-04", "expireDate": "2026-12-31", "fax": "-", "headerNote": "", "items": [{"desc": "ค่า Route Tracking Dashboard & Report System ( 45 คัน )", "price": 30000, "qty": 1}], "no": "Q-SCO-2026012", "notes": ["Report และ Dashboard สำหรับตรวจสอบการวิ่งรถของผู้รับเหมายกคันภายได้การดำเนินงานของ BRS โดยการ API ข้อมูลจากระบบ GPS ของผู้รับเหมา", "สามารถ Set Master Route พิกัดแต่ละเส้นทางได้", "สามารถ Update/Upload ข้อมูล Route plan และข้อมูลทะเบียนรถ และพชร. ที่มีการเปลี่ยนแปลงได้รายวัน", "แสดงข้อมูลพิกัดการวิ่งรถของผู้รับเหมาได้แบบ Real-time", "สามารถแสดงประวัติการเดินทางย้อนหลังได้", "สามารถแจ้ง Estimated Time Arrival (ETA) แต่ละจุดหมายของรถแต่ละคันได้", "ทำ Geofencing เพื่อการแจ้งเดือน กรณีวิ่งออกนอก Route ที่กำหนด", "ระบบการแจ้งเตือนรถเข้า - ออก จากโรงงาน ระบบการแจ้งเตือนรถเข้า - ออก จุดชั่งผ่าน รง.4", "อัตราค่าบริการตามใบเสนอราคาครอบคลุมจำนวนรถสูงสุด 45 คัน ในกรณีที่มีการใช้บริการเกินกว่าจำนวนดังกล่าว บริษัทขอคิดค่าบริการเพิ่มเติม ในอัตราคันละ 500 บาท/เดือน หากใช้บริการไม่เต็มเดือนคิดค่าบริการเต็มเดือน"], "preparer": "นาวิน มูลประเสริฐ", "priceUnit": "บาท/เดือน", "seq": 12, "subtotal": 30000, "taxId": "", "tel": "0870076004", "total": 30000, "updatedAt": "2026-06-04T09:00:00.000Z", "vat": false, "vatAmount": 0, "year": 2026}}], "numbers": [{"at": "2026-10-31T08:42:08.856Z", "no": "Q-SCO-2026013", "quoteId": "a0jlqasr7l81e53hzbon", "seq": 13, "year": 2026}]};

function importOldData() {
  setup();
  if (!OLD_DATA) { Logger.log('ไม่มีข้อมูลเดิมให้นำเข้า'); return; }
  withLock_(function () {
    var have = {}; readQuotes_().forEach(function (q) { have[q.id] = 1; });
    var nums = {}; readNumbers_().forEach(function (n) { nums[n.no] = 1; });
    var added = 0;
    OLD_DATA.quotes.forEach(function (q) {
      if (have[q.id]) return;
      writeQuote_(q.id, q.data, null); added++;
      if (q.data.no && !nums[q.data.no]) {
        var ns = sheet_(SHEET_NUMBERS);
        ns.getRange(ns.getLastRow() + 1, 1, 1, N_HEAD.length).setValues([[q.data.no, q.data.year || '', q.data.seq || '', q.id, q.data.createdAt || '']]);
        nums[q.data.no] = 1;
      }
    });
    (OLD_DATA.numbers || []).forEach(function (n) {
      if (nums[n.no]) return;
      var ns = sheet_(SHEET_NUMBERS);
      ns.getRange(ns.getLastRow() + 1, 1, 1, N_HEAD.length).setValues([[n.no, n.year || '', n.seq || '', n.quoteId || '', n.at || '']]);
      nums[n.no] = 1;
    });
    Logger.log('นำเข้าใบเสนอราคาเดิม ' + added + ' ใบ');
  });
}
