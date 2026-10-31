/**
 * ระบบบันทึกสุขภาพและกิจวัตรประจำวันนักเรียน (เวอร์ชัน 2)
 * - เก็บข้อมูลแยกตามปีการศึกษา/เดือน ย้อนดูและพิมพ์ซ้ำได้
 * - รองรับหลายห้องเรียนในไฟล์เดียว (อ.1 - ป.6) จำนวนนักเรียนไม่จำกัด
 * - นำเข้ารายชื่อจาก CSV/Excel/วางจาก Excel
 * - ตั้งค่าโรงเรียน/ครู/ผอ. ได้จากหน้า "ตั้งค่า" ในแอป
 * - (เวอร์ชัน 3) ระบบหลายผู้ใช้: เข้าสู่ระบบด้วย Username + PIN 6 หลัก, สิทธิ์รายห้อง,
 *   ผู้ดูแลระบบ (Admin) ยืนยันด้วยอีเมล Google, ผู้อำนวยการดูได้อย่างเดียว
 *
 * ชีตที่ระบบใช้ (สร้างให้อัตโนมัติเมื่อเปิดแอปครั้งแรก):
 *   ตั้งค่า, รายชื่อนักเรียน, วันหยุดนักขัตฤกษ์, เกณฑ์โภชนาการ,
 *   ข้อมูลกิจวัตร, ข้อมูลน้ำหนักส่วนสูง, ข้อมูลสุขภาพรายบุคคล, ผู้ใช้งาน
 */

var SH = {
  settings: 'ตั้งค่า',
  roster: 'รายชื่อนักเรียน',
  routine: 'ข้อมูลกิจวัตร',
  health: 'ข้อมูลน้ำหนักส่วนสูง',
  check: 'ข้อมูลตรวจสุขภาพ',
  indiv: 'ข้อมูลสุขภาพรายบุคคล',
  holidays: 'วันหยุดนักขัตฤกษ์',
  criteria: 'เกณฑ์โภชนาการ',
  rooms: 'ห้องเรียน',
  users: 'ผู้ใช้งาน'
};

var THAI_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
                   'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
var THAI_DOW = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

// [key, ป้ายในชีต, ค่าเริ่มต้น]
var SETTING_DEFS = [
  ['schoolName', 'ชื่อโรงเรียน', ''],
  ['grade', 'ระดับชั้น / ห้อง', ''],
  ['academicYear', 'ปีการศึกษา (พ.ศ.)', 2569],
  ['startMonth', 'เดือนที่เริ่มปีการศึกษา (1-12)', 5],
  ['currentMonth', 'เดือนที่บันทึกอยู่ปัจจุบัน (1-12)', new Date().getMonth() + 1],
  ['teacherName', 'ชื่อ-สกุล ครูประจำชั้น', ''],
  ['teacherPos', 'ตำแหน่งครูประจำชั้น', 'ครูประจำชั้น'],
  ['directorName', 'ชื่อ-สกุล ผู้อำนวยการ', ''],
  ['directorPos', 'ตำแหน่งผู้อำนวยการ', 'ผู้อำนวยการโรงเรียน']
];


/* ------------------------------------------------------------------ */
/*  เมนูในไฟล์ Google Sheet + หน้าเว็บ                                   */
/*  เมนูทุกตัวใช้ได้เฉพาะ "ผู้ดูแลระบบ (Admin)" ที่เปิดจากไฟล์ Sheet เท่านั้น    */
/*  (ฟังก์ชันสาธารณะใน Apps Script ถูกเรียกจากหน้าเว็บได้ จึงต้องมี menuGuard_) */
/* ------------------------------------------------------------------ */

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('ระบบบันทึกสุขภาพ')
      .addItem('1) ตั้งค่าความปลอดภัย / ผู้ดูแลระบบ (ทำครั้งแรก)', 'setupSecurity')
      .addItem('เพิ่มอีเมลผู้ดูแลระบบ (Admin)', 'manageAdmins')
      .addSeparator()
      .addItem('ตรวจ/สร้างโครงสร้างชีต', 'setupSheets')
      .addItem('ล้างแคชระบบ (ใช้หลังแก้ชีตด้วยมือ)', 'clearCacheMenu')
      .addItem('ลบชีตเก่าที่ไม่ใช้แล้ว', 'cleanupOldSheets')
      .addItem('ล้างข้อมูลทดสอบทั้งหมด (เริ่มใหม่)', 'clearAllTestData')
      .addToUi();
  } catch (e) { /* ไม่มีหน้าจอ UI (เช่น เรียกผ่านเว็บ) */ }
}

// คืนอ็อบเจกต์ UI เมื่อเรียกจากเมนูในไฟล์ Sheet โดย Admin เท่านั้น มิฉะนั้นคืน null
// - getUi() ใช้ได้เฉพาะตอนเปิดไฟล์ Sheet (เรียกผ่านเว็บแอป/google.script.run จะ error ทันที)
// - อีเมลผู้ใช้ปัจจุบันที่ Google ยืนยันแล้วต้องอยู่ในรายชื่อ Admin
function menuGuard_(allowFirstSetup) {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { return null; }
  var admins = getAdminEmails_();
  if (!admins.length) {
    if (allowFirstSetup) return ui;
    ui.alert('ยังไม่ได้ตั้งผู้ดูแลระบบ', 'กรุณาเลือกเมนู "1) ตั้งค่าความปลอดภัย / ผู้ดูแลระบบ" ก่อน', ui.ButtonSet.OK);
    return null;
  }
  if (admins.indexOf(menuEmail_()) < 0) {
    ui.alert('ไม่อนุญาต', 'เมนูนี้ใช้ได้เฉพาะผู้ดูแลระบบ (Admin) เท่านั้น', ui.ButtonSet.OK);
    return null;
  }
  return ui;
}

function menuEmail_() {
  var e = '';
  try { e = String(Session.getActiveUser().getEmail() || ''); } catch (x) {}
  if (!e) { try { e = String(Session.getEffectiveUser().getEmail() || ''); } catch (x2) {} }
  return e.trim().toLowerCase();
}

// ทำครั้งแรก: สร้างกุญแจเข้ารหัส PIN, สร้างชีต "ผู้ใช้งาน" และตั้งผู้ดูแลระบบคนแรก = อีเมลของผู้ที่เปิดเมนูนี้
function setupSecurity() {
  var ui = menuGuard_(true);
  if (!ui) return;
  withLock_(function () { ensurePepper_(); ensureSetup_(); });
  var admins = getAdminEmails_();
  if (!admins.length) {
    var me = menuEmail_();
    if (!me) {
      var r = ui.prompt('อีเมลผู้ดูแลระบบ', 'ระบบอ่านอีเมลของคุณไม่ได้ กรุณาพิมพ์อีเมล Google ของผู้ดูแลระบบ', ui.ButtonSet.OK_CANCEL);
      if (r.getSelectedButton() !== ui.Button.OK) return;
      me = String(r.getResponseText() || '').trim().toLowerCase();
    }
    if (!validEmail_(me)) { ui.alert('อีเมลไม่ถูกต้อง', 'ยกเลิกการตั้งค่า', ui.ButtonSet.OK); return; }
    setAdminEmails_([me]);
    admins = [me];
  }
  ui.alert('ตั้งค่าความปลอดภัยเรียบร้อย',
    'ผู้ดูแลระบบ (Admin): ' + admins.join(', ') +
    '\n\nสร้างชีต "ผู้ใช้งาน" แล้ว (เก็บเฉพาะ PIN ที่เข้ารหัส ห้ามแก้ด้วยมือ)' +
    '\nต่อไป: Deploy เว็บแอป แล้วเข้าด้วยลิงก์ของ Admin เพื่อสร้างบัญชีครู/ผอ. ในแท็บ "ผู้ใช้งาน"', ui.ButtonSet.OK);
}

function manageAdmins() {
  var ui = menuGuard_(false);
  if (!ui) return;
  var admins = getAdminEmails_();
  var r = ui.prompt('ผู้ดูแลระบบ (Admin)',
    'ปัจจุบัน: ' + admins.join(', ') + '\n\nพิมพ์อีเมล Google ที่ต้องการเพิ่มเป็น Admin (เว้นว่างแล้วกดยกเลิกถ้าไม่ต้องการเพิ่ม)\n' +
    '(ลบ Admin: Project Settings > Script properties > ADMIN_EMAILS)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var em = String(r.getResponseText() || '').trim().toLowerCase();
  if (!em) return;
  if (!validEmail_(em)) { ui.alert('อีเมลไม่ถูกต้อง'); return; }
  if (admins.indexOf(em) < 0) { admins.push(em); setAdminEmails_(admins); }
  ui.alert('บันทึกแล้ว', 'ผู้ดูแลระบบ: ' + admins.join(', ') +
    '\n\nหมายเหตุ: Admin ต้องเข้าเว็บแอปด้วยลิงก์ของ Admin (Deploy แบบ "ผู้ใช้ที่เข้าถึงเว็บแอป") และมีสิทธิ์แก้ไขไฟล์ Sheet นี้', ui.ButtonSet.OK);
}

function setupSheets() {
  var ui = menuGuard_(false);
  if (!ui) return;
  ensureSetup_(true);   // บังคับตรวจเต็มรูปแบบ (ไม่ใช้ธงแคช)
  ui.alert('ตรวจและสร้างโครงสร้างชีตเรียบร้อยแล้ว');
}

// ล้างแคชทั้งหมดทันที (ข้อมูลต้นฉบับในชีตไม่ถูกแตะ) — ใช้เมื่อแก้ข้อมูลในชีตด้วยมือแล้วอยากให้แอปเห็นทันที
function clearCacheMenu() {
  var ui = menuGuard_(false);
  if (!ui) return;
  bump_();
  try { CacheService.getScriptCache().remove('SETUP_OK:' + SETUP_VERSION); } catch (e) {}
  ui.alert('ล้างแคชเรียบร้อยแล้ว', 'การเปิดแอปครั้งถัดไปจะอ่านข้อมูลล่าสุดจากชีตทั้งหมด (อาจช้ากว่าปกติ 1 ครั้ง)', ui.ButtonSet.OK);
}

// ชีตของระบบเวอร์ชันเก่า ที่เวอร์ชันนี้ไม่ใช้แล้ว
var OLD_SHEETS = ['ชั่งน้ำหนัก-วัดส่วนสูง', 'บันทึกการดื่มนม', 'บันทึกการแปรงฟัน',
                  'บันทึกสุขภาพรายบุคคล', 'บันทึกสุขภาพรายบุคคล_Data'];

function cleanupOldSheets() {
  var ui = menuGuard_(false);
  if (!ui) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSetup_();
  var found = OLD_SHEETS.filter(function (n) { return ss.getSheetByName(n); });
  if (!found.length) { ui.alert('ไม่พบชีตเก่าที่ต้องลบ'); return; }
  var ans = ui.alert('ลบชีตเก่าที่ระบบเวอร์ชันนี้ไม่ใช้แล้ว',
    'จะลบชีตต่อไปนี้ (ข้อมูลในชีตเหล่านี้จะหายถาวร):\n\n- ' + found.join('\n- ') +
    '\n\nข้อมูลที่แอปบันทึกไว้อยู่ในชีต "ข้อมูล..." ซึ่งจะไม่ถูกลบ ต้องการลบหรือไม่?', ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  found.forEach(function (n) { ss.deleteSheet(ss.getSheetByName(n)); });
  ui.alert('ลบแล้ว ' + found.length + ' ชีต');
}

// ล้างรายชื่อนักเรียนและข้อมูลบันทึกทั้งหมด (นม แปรงฟัน น้ำหนัก-ส่วนสูง ประวัติสุขภาพ) เก็บการตั้งค่า/ห้อง/วันหยุด/ตารางเกณฑ์/ผู้ใช้งานไว้
function clearAllTestData() {
  var ui = menuGuard_(false);
  if (!ui) return;
  var ans = ui.alert('ล้างข้อมูลทั้งหมดเพื่อเริ่มใหม่',
    'จะลบ: รายชื่อนักเรียนทุกห้อง, ข้อมูลดื่มนม, แปรงฟัน, น้ำหนัก-ส่วนสูง, ตรวจสุขภาพรายสัปดาห์, ประวัติสุขภาพรายบุคคล\n' +
    'จะเก็บไว้: การตั้งค่า, ห้องเรียนและชื่อครู, วันหยุด, ตารางเกณฑ์โภชนาการ, บัญชีผู้ใช้งาน\n\nลบแล้วกู้คืนไม่ได้ ต้องการดำเนินการหรือไม่?', ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  withLock_(function () {
    var o = ensureSetup_();
    [o.routine, o.health, o.check, o.indiv].forEach(function (sh) {
      var last = sh.getLastRow();
      if (last > 1) sh.getRange(2, 1, last - 1, sh.getMaxColumns()).clearContent();
    });
    var h = rosterHeaderRow_(o.roster), rl = o.roster.getLastRow();
    if (rl > h) o.roster.getRange(h + 1, 1, rl - h, 8).clearContent();
  }, ['roster', 'data']);
  ui.alert('ล้างข้อมูลเรียบร้อยแล้ว เริ่มนำเข้ารายชื่อใหม่ได้เลย');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบบันทึกสุขภาพและกิจวัตรประจำวันนักเรียน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/* ================================================================== */
/*  ระบบผู้ใช้งาน / เข้าสู่ระบบ / สิทธิ์ (เพิ่มใหม่ในเวอร์ชันนี้)               */
/*                                                                      */
/*  หลักการ                                                              */
/*  1) ฟังก์ชันสาธารณะที่หน้าเว็บเรียกได้มีทางเดียวคือ rpc() เท่านั้น          */
/*     ฟังก์ชันงานเดิมทั้งหมดถูกเปลี่ยนเป็นชื่อลงท้าย "_" (เรียกจากหน้าเว็บตรง ๆ ไม่ได้) */
/*  2) rpc() ตรวจ Session + บทบาท + ห้องที่ได้รับสิทธิ์ ก่อนเรียกฟังก์ชันงานทุกครั้ง */
/*  3) PIN เก็บเป็นค่าแฮช (HMAC-SHA256 + salt รายคน + กุญแจลับใน Script Properties) */
/*  4) Admin ยืนยันตัวตนด้วยอีเมล Google ที่ Google ยืนยันแล้ว                  */
/*     (Session.getActiveUser) — ไม่เคยเชื่ออีเมลที่ส่งมาจากหน้าเว็บ              */
/* ================================================================== */

var ROLE_ADMIN = 'admin', ROLE_TEACHER = 'teacher', ROLE_DIRECTOR = 'director';
var R_ALL = [ROLE_ADMIN, ROLE_TEACHER, ROLE_DIRECTOR];
var R_WRITE = [ROLE_ADMIN, ROLE_TEACHER];   // ผู้ที่บันทึกข้อมูลได้ (ผอ. อ่านอย่างเดียว)
var R_ADMIN = [ROLE_ADMIN];

var AUTH = {
  maxFails: 5,            // กรอก PIN ผิดติดต่อกันกี่ครั้งแล้วล็อกบัญชี
  idleSec: 30 * 60,       // ไม่มีการใช้งานกี่วินาทีแล้ว Session หมดอายุ
  maxSec: 6 * 60 * 60,    // อายุ Session สูงสุด (ขีดจำกัดของ CacheService = 6 ชม.)
  hashRounds: 200         // จำนวนรอบแฮช PIN
};

var USER_HEADERS = ['Username', 'ชื่อที่แสดง', 'บทบาท (teacher/director)', 'ห้องที่เข้าถึงได้ (JSON)', 'PIN (แฮช ห้ามแก้)',
                    'Salt', 'สถานะ (active/disabled/locked)', 'PIN ผิดติดต่อกัน', 'เข้าสู่ระบบล่าสุด', 'รอบเซสชัน', 'สร้างเมื่อ', 'แก้ไขเมื่อ'];
var UCOL = { user: 1, name: 2, role: 3, rooms: 4, hash: 5, salt: 6, status: 7, fails: 8, last: 9, ver: 10, created: 11, updated: 12 };

/* ---------- ตัวช่วยทั่วไป ---------- */

function now_() { return new Date().getTime(); }
function stamp_() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
function validEmail_(e) { return /^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(String(e || '')); }
function randHex_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); }
function bytesToHex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function safeEq_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function userErr_(msg, code) { var e = new Error(msg); e.appCode = code || 'DENIED'; return e; }
function fail_(msg, code) { throw userErr_(msg, code); }
function ok_(data) { return { ok: true, data: data }; }
function clip_(v, n) { return String(v == null ? '' : v).substring(0, n); }
function plainObj_(v) { return (v && typeof v === 'object' && !Array.isArray(v)) ? v : {}; }
function intIn_(v, min, max, label) {
  var n = Number(v);
  if (v === '' || v === null || !isFinite(n) || Math.floor(n) !== n || n < min || n > max) fail_(label + ' ไม่ถูกต้อง', 'BAD');
  return n;
}
function strArr_(v, label, max) {
  if (!Array.isArray(v) || v.length > max) fail_('ข้อมูล' + label + 'ไม่ถูกต้อง', 'BAD');
  return v.map(function (x) { return String(x == null ? '' : x).trim(); });
}
function activeEmail_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { return ''; }
}

/* ---------- ผู้ดูแลระบบ (Admin) + กุญแจลับ ---------- */

function getAdminEmails_() {
  var raw = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || '';
  return raw.split(',').map(function (x) { return x.trim().toLowerCase(); }).filter(function (x) { return x; });
}
function setAdminEmails_(list) {
  PropertiesService.getScriptProperties().setProperty('ADMIN_EMAILS', list.join(','));
}
function ensurePepper_() {
  var p = PropertiesService.getScriptProperties(), v = p.getProperty('PIN_PEPPER');
  if (!v) { v = randHex_(); p.setProperty('PIN_PEPPER', v); }
  return v;
}

/* ---------- แฮช PIN ---------- */

function hashPin_(pin, salt, username) {
  var pepper = ensurePepper_();
  var h = Utilities.computeHmacSha256Signature(salt + ':' + username + ':' + pin, pepper);
  for (var i = 0; i < AUTH.hashRounds; i++) {
    h = Utilities.computeHmacSha256Signature(Utilities.base64Encode(h) + ':' + salt, pepper);
  }
  return 'h1$' + bytesToHex_(h);
}
function pinProblem_(pin) {
  pin = String(pin == null ? '' : pin);
  if (!/^[0-9]{6}$/.test(pin)) return 'PIN ต้องเป็นตัวเลข 6 หลักพอดี';
  if (/^(\d)\1{5}$/.test(pin)) return 'PIN ง่ายเกินไป (เลขซ้ำกันทั้ง 6 หลัก)';
  if ('0123456789'.indexOf(pin) >= 0 || '9876543210'.indexOf(pin) >= 0) return 'PIN ง่ายเกินไป (เลขเรียงกัน)';
  return '';
}

/* ---------- ชีตผู้ใช้งาน ---------- */

function ensureUsers_(ss) {
  var sh = ss.getSheetByName(SH.users);
  if (!sh) {
    sh = ss.insertSheet(SH.users);
    sh.getRange(1, 1, sh.getMaxRows(), USER_HEADERS.length).setNumberFormat('@'); // ทุกช่องเป็นข้อความ (กัน Sheets แปลงเป็นสูตร/ตัวเลข)
    sh.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    try { sh.protect().setWarningOnly(true).setDescription('ห้ามแก้ไขด้วยมือ — จัดการผู้ใช้จากแท็บ "ผู้ใช้งาน" ในเว็บแอป'); } catch (e) {}
  }
  return sh;
}
function ensureRows_(sh, needRows) {
  var m = sh.getMaxRows();
  if (m < needRows) sh.insertRowsAfter(m, needRows - m + 200);
}
function parseRooms_(v) {
  try { var a = JSON.parse(String(v == null || v === '' ? '[]' : v)); return Array.isArray(a) ? a.map(String) : []; }
  catch (e) { return []; }
}
function loadUsers_() {
  var sh = ensureUsers_(SpreadsheetApp.getActiveSpreadsheet()), last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, USER_HEADERS.length).getValues().forEach(function (r, i) {
    var u = String(r[0]).trim().toLowerCase();
    if (!u) return;
    out.push({ row: i + 2, username: u, name: String(r[1]), role: String(r[2]).trim(), rooms: parseRooms_(r[3]),
               hash: String(r[4]), salt: String(r[5]), status: String(r[6] || 'active').trim(),
               fails: Number(r[7]) || 0, lastLogin: String(r[8]), ver: Number(r[9]) || 0,
               created: String(r[10]), updated: String(r[11]) });
  });
  return out;
}
// username: ชื่อผู้ใช้ (พิมพ์เล็ก) / fresh=true = อ่านจากชีตจริง (ใช้ตอนเข้าสู่ระบบ/แก้ไข/ลบบัญชี เพื่อให้ตัวนับ PIN ผิดแม่นยำเสมอ)
function findUser_(username, fresh) {
  var l = readUsers_(fresh);
  for (var i = 0; i < l.length; i++) if (l[i].username === username) return l[i];
  return null;
}
function userPublic_(u) {
  return { username: u.username, name: u.name, role: u.role, rooms: u.rooms, status: u.status,
           fails: u.fails, lastLogin: u.lastLogin, created: u.created, updated: u.updated };
}
function usersPayload_(fresh) {
  return { users: readUsers_(fresh).map(userPublic_), rooms: getRooms_(fresh).map(function (r) { return { name: r.name, full: r.full }; }),
           maxFails: AUTH.maxFails };
}

// เมื่อเปลี่ยนชื่อ/ลบห้อง ให้สิทธิ์ห้องของผู้ใช้ตามไปด้วย (เรียกภายใน withLock_ ของผู้เรียกเท่านั้น ห้ามล็อกซ้อน)
// map: { ชื่อห้องเดิม : ชื่อห้องใหม่ | null(ถูกลบ) } ห้องที่ไม่อยู่ใน map = ไม่เปลี่ยน
function remapUserRooms_(map) {
  var sh = ensureUsers_(SpreadsheetApp.getActiveSpreadsheet()), last = sh.getLastRow();
  if (last < 2) return;
  var rg = sh.getRange(2, UCOL.rooms, last - 1, 1), changed = false;
  var out = rg.getValues().map(function (r) {
    var list = parseRooms_(r[0]), nl = [];
    list.forEach(function (n) {
      var m = Object.prototype.hasOwnProperty.call(map, n) ? map[n] : n;
      if (m && nl.indexOf(m) < 0) nl.push(m);
    });
    var s = JSON.stringify(nl);
    if (s !== JSON.stringify(list)) changed = true;
    return [s];
  });
  if (changed) rg.setValues(out);
}

/* ---------- Session (เก็บฝั่งเซิร์ฟเวอร์ใน CacheService) ---------- */

function sessKey_(token) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  return 'S1:' + Utilities.base64EncodeWebSafe(d);
}
function newSession_(data) {
  var token = randHex_();
  data.iat = now_(); data.last = data.iat;
  CacheService.getScriptCache().put(sessKey_(token), JSON.stringify(data), AUTH.idleSec);
  return token;
}
function killSession_(token) {
  if (typeof token === 'string' && /^[0-9a-f]{64}$/.test(token)) CacheService.getScriptCache().remove(sessKey_(token));
}
function readSession_(token) {
  if (typeof token !== 'string' || !/^[0-9a-f]{64}$/.test(token)) return null;
  var cache = CacheService.getScriptCache(), key = sessKey_(token), raw = cache.get(key);
  if (!raw) return null;
  var s;
  try { s = JSON.parse(raw); } catch (e) { cache.remove(key); return null; }
  var t = now_();
  if (t - s.last > AUTH.idleSec * 1000 || t - s.iat > AUTH.maxSec * 1000) { cache.remove(key); return null; }
  if (t - s.last > 30000) { // ต่ออายุแบบเลื่อนตามการใช้งาน (ไม่เกินอายุสูงสุด)
    s.last = t;
    var ttl = Math.min(AUTH.idleSec, Math.max(1, Math.floor((AUTH.maxSec * 1000 - (t - s.iat)) / 1000)));
    cache.put(key, JSON.stringify(s), ttl);
  }
  return s;
}

// ตรวจ Session ทุกครั้งที่เรียก และอ่านสิทธิ์ล่าสุดของผู้ใช้จากชีต (ปิดบัญชี/เปลี่ยนสิทธิ์/รีเซ็ต PIN มีผลทันที)
function authenticate_(token) {
  var s = readSession_(token);
  if (!s) fail_('หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่', 'AUTH');
  if (s.role === ROLE_ADMIN) {
    var email = activeEmail_();
    if (!email || email !== s.email || getAdminEmails_().indexOf(email) < 0) { killSession_(token); fail_('ยืนยันตัวตนผู้ดูแลระบบไม่ผ่าน กรุณาเข้าสู่ระบบใหม่', 'AUTH'); }
    return { role: ROLE_ADMIN, username: email, name: 'ผู้ดูแลระบบ', email: email, allRooms: true, rooms: [], readOnly: false };
  }
  var u = findUser_(s.u);
  if (!u || u.status !== 'active' || u.ver !== s.v || u.role !== s.role ||
      (u.role !== ROLE_TEACHER && u.role !== ROLE_DIRECTOR)) {
    killSession_(token);
    fail_('บัญชีถูกปิดใช้งานหรือมีการเปลี่ยนแปลงสิทธิ์ กรุณาเข้าสู่ระบบใหม่', 'AUTH');
  }
  return { role: u.role, username: u.username, name: u.name, allRooms: u.role === ROLE_DIRECTOR,
           rooms: u.role === ROLE_TEACHER ? u.rooms : [], readOnly: u.role === ROLE_DIRECTOR };
}
function meOf_(ctx) {
  return { username: ctx.username, name: ctx.name, role: ctx.role, readOnly: ctx.readOnly, allRooms: ctx.allRooms, email: ctx.email || '' };
}

/* ---------- เข้าสู่ระบบ ---------- */

var MSG_BAD_LOGIN = 'Username หรือ PIN ไม่ถูกต้อง (กรอก PIN ผิดหลายครั้ง บัญชีจะถูกล็อก)';

function helloInfo_() {
  var email = activeEmail_(), admins = getAdminEmails_(), isAdmin = !!email && admins.indexOf(email) >= 0;
  return { google: { isAdmin: isAdmin, email: isAdmin ? email : '' }, adminConfigured: admins.length > 0, idleSec: AUTH.idleSec };
}

function loginRpc_(username, pin) {
  username = String(username == null ? '' : username).trim().toLowerCase();
  pin = String(pin == null ? '' : pin);
  var bad = { ok: false, code: 'BAD', message: MSG_BAD_LOGIN };
  if (!/^[a-z0-9._-]{3,30}$/.test(username) || !/^[0-9]{6}$/.test(pin)) {
    hashPin_('000000', 'x', 'x'); // ใช้เวลาใกล้เคียงกับกรณีปกติ
    return bad;
  }
  return withLock_(function () {
    var u = findUser_(username, true);   // อ่านจากชีตจริงเสมอ: ตัวนับ PIN ผิดและสถานะล็อกต้องแม่นยำ ไม่ผ่านแคช
    if (!u) { hashPin_(pin, 'x', username); return bad; }
    var good = safeEq_(hashPin_(pin, u.salt, u.username), u.hash);
    var sh = ensureUsers_(SpreadsheetApp.getActiveSpreadsheet());
    if (u.status !== 'active') { // บอกสถานะเฉพาะเมื่อ PIN ถูก (กันคนนอกเดาว่ามี Username นี้อยู่)
      if (!good) return bad;
      return { ok: false, code: 'DISABLED', message: u.status === 'locked'
        ? 'บัญชีนี้ถูกล็อกเพราะกรอก PIN ผิดหลายครั้ง กรุณาติดต่อผู้ดูแลระบบเพื่อปลดล็อก'
        : 'บัญชีนี้ถูกปิดใช้งาน กรุณาติดต่อผู้ดูแลระบบ' };
    }
    if (!good) {
      var f = u.fails + 1, st = f >= AUTH.maxFails ? 'locked' : 'active';
      sh.getRange(u.row, UCOL.status, 1, 2).setValues([[st, f]]);
      return bad;
    }
    if (u.role !== ROLE_TEACHER && u.role !== ROLE_DIRECTOR) return bad;
    sh.getRange(u.row, UCOL.fails, 1, 2).setValues([[0, stamp_()]]);
    var token = newSession_({ role: u.role, u: u.username, v: u.ver });
    var ctx = { username: u.username, name: u.name, role: u.role, readOnly: u.role === ROLE_DIRECTOR, allRooms: u.role === ROLE_DIRECTOR };
    return ok_({ token: token, me: meOf_(ctx), idleSec: AUTH.idleSec });
  }, ['users']);
}

function adminEnterRpc_() {
  var email = activeEmail_();
  if (!email || getAdminEmails_().indexOf(email) < 0) {
    return { ok: false, code: 'FORBIDDEN', message: 'บัญชี Google ที่ใช้อยู่ไม่ใช่ผู้ดูแลระบบ (ต้องเปิดจากลิงก์ของ Admin และเข้าด้วยอีเมลที่ตั้งเป็น Admin)' };
  }
  var token = newSession_({ role: ROLE_ADMIN, email: email });
  return ok_({ token: token, me: meOf_({ username: email, name: 'ผู้ดูแลระบบ', role: ROLE_ADMIN, readOnly: false, allRooms: true, email: email }), idleSec: AUTH.idleSec });
}

/* ---------- ขอบเขตข้อมูลตามสิทธิ์ ---------- */

function roomAllowed_(ctx, room) { return ctx.allRooms || ctx.rooms.indexOf(String(room)) >= 0; }
// คืน null = เข้าถึงได้ทุกคน (Admin/ผอ.) ; ครู = แผนที่ id นักเรียนในห้องที่ได้รับสิทธิ์
function allowedIdMap_(ctx) {
  if (ctx.allRooms) return null;
  var m = {};
  readRosterRaw_().forEach(function (s) { if (ctx.rooms.indexOf(s.room) >= 0) m[s.id] = 1; });
  return m;
}
function assertIds_(ctx, ids) {
  var m = allowedIdMap_(ctx);
  if (!m) return;
  ids.forEach(function (id) { if (!m[id]) fail_('ไม่มีสิทธิ์เข้าถึงข้อมูลนักเรียนนอกห้องที่ได้รับสิทธิ์', 'FORBIDDEN'); });
}
function assertId_(ctx, id) {
  id = String(id == null ? '' : id).trim();
  if (!id || id === 'undefined' || id === 'null') fail_('ไม่ได้เลือกนักเรียน', 'BAD');
  assertIds_(ctx, [id]);
  return id;
}
function pick_(obj, idMap) {
  var out = {};
  Object.keys(obj || {}).forEach(function (k) { if (idMap[k]) out[k] = obj[k]; });
  return out;
}
function scopeResult_(ctx, r) {
  if (!r || ctx.allRooms) return r;
  if (r.students) r.students = r.students.filter(function (s) { return ctx.rooms.indexOf(s.room) >= 0; });
  if (r.rooms) r.rooms = r.rooms.filter(function (x) { return ctx.rooms.indexOf(x.name) >= 0; });
  if (r.revs) {
    var rv = {};
    Object.keys(r.revs).forEach(function (k) { if (ctx.rooms.indexOf(k) >= 0) rv[k] = r.revs[k]; });
    r.revs = rv;
  }
  return r;
}

/* ---------- เลขรุ่นรายชื่อรายห้อง (กันบันทึกรายชื่อทับกันเมื่อหลายคนแก้พร้อมกัน) ---------- */
function shortHash_(s) {
  return bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)).substring(0, 16);
}
// list = ผลของ readRosterRaw_, rooms = ผลของ getRooms_ => { ชื่อห้อง: เลขรุ่น } (ห้องที่ไม่มีนักเรียนก็มีเลขรุ่น)
function roomRevs_(list, rooms) {
  var by = {};
  list.forEach(function (s) { (by[s.room] = by[s.room] || []).push([s.id, s.name, s.nickname, s.gender, s.birth, s.note]); });
  var out = {};
  rooms.forEach(function (r) { out[r.name] = shortHash_(JSON.stringify(by[r.name] || [])); });
  return out;
}

// ปีการศึกษา/เดือนปัจจุบันตามเวลาไทย (ตรงกับ todayPeriod ในหน้าเว็บ) — ใช้ส่งข้อมูลเดือนแรกไปพร้อมกับการเข้าสู่ระบบ
function todayPeriod_(st) {
  var tz = 'Asia/Bangkok', now = new Date();
  var y = Number(Utilities.formatDate(now, tz, 'yyyy')), m = Number(Utilities.formatDate(now, tz, 'M'));
  var sm = Number(st.startMonth) || 5;
  return { year: (m >= sm ? y : y - 1) + 543, month: m };
}

function getBootstrapFor_(ctx, wantPeriod) {
  var b = getBootstrap_();
  b.me = meOf_(ctx);
  b.revs = roomRevs_(readRosterRaw_(), getRooms_());
  if (!ctx.allRooms) b = scopeResult_(ctx, b);
  var p = todayPeriod_(b.settings);
  b.p0 = p;
  if (wantPeriod) b.period = getPeriodFor_(ctx, p.year, p.month);   // ส่งข้อมูลเดือนแรกไปพร้อมกัน ลดการเรียกเซิร์ฟเวอร์ 1 รอบตอนเข้าใช้งาน
  return b;
}

/* ---------- ตัวห่อฟังก์ชันงาน: ตรวจสิทธิ์ + ตรวจข้อมูลที่ส่งมา แล้วเรียกฟังก์ชันงานเดิม ---------- */

function getPeriodFor_(ctx, year, month) {
  year = intIn_(year, 2400, 2800, 'ปีการศึกษา'); month = intIn_(month, 1, 12, 'เดือน');
  var d = getPeriod_(year, month);
  var m = allowedIdMap_(ctx);
  if (!m) return d;
  return { milk: pick_(d.milk, m), brush: pick_(d.brush, m), health: pick_(d.health, m), ref: pick_(d.ref, m), check: pick_(d.check, m) };
}
function getGrowthFor_(ctx, year, id) {
  return getGrowth_(intIn_(year, 2400, 2800, 'ปีการศึกษา'), assertId_(ctx, id));
}
function getIndividualFor_(ctx, id) { return getIndividual_(assertId_(ctx, id)); }

function idSet_(ids) { var m = {}; ids.forEach(function (i) { m[i] = 1; }); return m; }

function saveRoutineFor_(ctx, year, month, type, grid, ids, mode) {
  year = intIn_(year, 2400, 2800, 'ปีการศึกษา'); month = intIn_(month, 1, 12, 'เดือน');
  ids = strArr_(ids, 'รายชื่อนักเรียน', 5000); // ต้องส่งมาเสมอ (กันการเขียนทับข้อมูลทั้งไฟล์)
  assertIds_(ctx, ids);
  mode = mode === 'diff' ? 'diff' : '';
  grid = plainObj_(grid);
  var set = idSet_(ids);
  var clean = {};
  Object.keys(grid).forEach(function (id) {
    if (!set[id]) fail_('ข้อมูลไม่ตรงกับนักเรียนที่เลือก', 'BAD');
    if (mode === 'diff') {
      var g = plainObj_(grid[id]);
      if ((Array.isArray(g.a) && g.a.length > 62) || (Array.isArray(g.r) && g.r.length > 62)) fail_('ข้อมูลวันที่ไม่ถูกต้อง', 'BAD');
      clean[id] = { a: daysList_(g.a), r: daysList_(g.r) };
      return;
    }
    var days = Array.isArray(grid[id]) ? grid[id] : [];
    if (days.length > 62) fail_('ข้อมูลวันที่ไม่ถูกต้อง', 'BAD');
    clean[id] = daysList_(days);
  });
  return saveRoutine_(year, month, type, clean, ids, mode);
}
function saveCheckFor_(ctx, year, month, map, ids, mode) {
  year = intIn_(year, 2400, 2800, 'ปีการศึกษา'); month = intIn_(month, 1, 12, 'เดือน');
  ids = strArr_(ids, 'รายชื่อนักเรียน', 5000);
  assertIds_(ctx, ids);
  mode = mode === 'diff' ? 'diff' : '';
  map = plainObj_(map);
  var set = idSet_(ids), clean = {};
  Object.keys(map).forEach(function (id) {
    if (!set[id]) fail_('ข้อมูลไม่ตรงกับนักเรียนที่เลือก', 'BAD');
    if (mode === 'diff') {
      var df = plainObj_(map[id]), o = {};
      Object.keys(df).forEach(function (k) {
        var p = Number(k), c = String(df[k]);
        if (!(p >= 0 && p < 20 && Math.floor(p) === p) || (c !== 'g' && c !== 'x' && c !== '-')) fail_('ข้อมูลตรวจสุขภาพไม่ถูกต้อง', 'BAD');
        o[p] = c;
      });
      clean[id] = o;
    } else clean[id] = map[id];
  });
  return saveCheck_(year, month, clean, ids, mode);
}
function saveHealthFor_(ctx, year, month, rows, ids, mode) {
  year = intIn_(year, 2400, 2800, 'ปีการศึกษา'); month = intIn_(month, 1, 12, 'เดือน');
  ids = strArr_(ids, 'รายชื่อนักเรียน', 5000);
  assertIds_(ctx, ids);
  mode = mode === 'diff' ? 'diff' : '';
  if (!Array.isArray(rows) || rows.length > 5000) fail_('ข้อมูลน้ำหนัก-ส่วนสูงไม่ถูกต้อง', 'BAD');
  var set = idSet_(ids);
  rows = rows.map(function (r) {
    r = plainObj_(r);
    var id = String(r.id == null ? '' : r.id).trim();
    if (!set[id]) fail_('ข้อมูลไม่ตรงกับนักเรียนที่เลือก', 'BAD');
    var o = { id: id };
    ['w', 'h'].forEach(function (f) {
      if (mode === 'diff' && !Object.prototype.hasOwnProperty.call(r, f)) return;   // โหมด diff: ส่งมาเฉพาะช่องที่เปลี่ยน
      o[f] = r[f];
    });
    var w = parseFloat(o.w);
    if (!isNaN(w) && (w < 5 || w > 150)) fail_('น้ำหนักต้องอยู่ระหว่าง 5-150 กก.', 'BAD');
    return o;
  });
  return saveHealth_(year, month, rows, ids, mode);
}
function saveIndividualFor_(ctx, id, data) {
  id = assertId_(ctx, id);
  data = plainObj_(data);
  if (JSON.stringify(data).length > 30000) fail_('ข้อมูลยาวเกินกำหนด', 'BAD');
  return saveIndividual_(id, data);
}
function clearIndividualFor_(ctx, id) { return clearIndividual_(assertId_(ctx, id)); }

function cleanStudent_(r) {
  r = plainObj_(r);
  return { id: clip_(r.id, 30).trim(), name: clip_(r.name, 150), nickname: clip_(r.nickname, 60),
           gender: clip_(r.gender, 10), birth: clip_(r.birth, 20), note: clip_(r.note, 200) };
}
// baseRev = "เลขรุ่นรายชื่อห้องนี้" ที่หน้าเว็บโหลดไว้ ถ้าไม่ตรงกับปัจจุบัน (มีคนอื่นแก้ไขรายชื่อ/ลบนักเรียนไปก่อน) จะไม่บันทึกทับ
function saveRosterFor_(ctx, rows, room, baseRev) {
  room = String(room == null ? '' : room).trim();
  if (!getRooms_(true).some(function (r) { return r.name === room; })) fail_('ไม่พบห้องเรียนนี้', 'BAD');
  if (!roomAllowed_(ctx, room)) fail_('ไม่มีสิทธิ์จัดการรายชื่อของห้องนี้', 'FORBIDDEN');
  if (!Array.isArray(rows) || rows.length > 1000) fail_('ข้อมูลรายชื่อไม่ถูกต้อง', 'BAD');
  return scopeResult_(ctx, saveRoster_(rows.map(cleanStudent_), room, clip_(baseRev, 40)));
}

/* ---------- ลบนักเรียน (ลบข้อมูลออกจากชีตจริง) ---------- */
// ขั้นที่ 1: ตรวจผลกระทบ — นับข้อมูลที่ผูกกับนักเรียนคนนี้ในทุกชีต (ยังไม่ลบอะไร)
function countStudentData_(id) {
  var o = ensureSetup_(), out = { milk: 0, brush: 0, health: 0, check: 0, indiv: 0 };
  var rl = o.routine.getLastRow();
  if (rl > 1) o.routine.getRange(2, 3, rl - 1, 2).getValues().forEach(function (r) {
    if (String(r[1]) === id && (r[0] === 'milk' || r[0] === 'brush')) out[r[0]]++;
  });
  [['health', o.health, 3], ['check', o.check, 3], ['indiv', o.indiv, 1]].forEach(function (d) {
    var last = d[1].getLastRow();
    if (last < 2) return;
    d[1].getRange(2, d[2], last - 1, 1).getValues().forEach(function (r) { if (String(r[0]) === id) out[d[0]]++; });
  });
  return out;
}

// ขั้นที่ 2: ลบแถวของนักเรียนคนนี้ออกจากชีตข้อมูลทั้งหมด (ต้องเรียกภายใน withLock_)
function removeStudentData_(id) {
  var o = ensureSetup_(), out = { milk: 0, brush: 0, health: 0, check: 0, indiv: 0 };
  var rl = o.routine.getLastRow();
  if (rl > 1) {
    var rows = [];
    o.routine.getRange(2, 3, rl - 1, 2).getValues().forEach(function (r, i) {
      if (String(r[1]) === id) { rows.push(i + 2); if (out[r[0]] !== undefined) out[r[0]]++; }
    });
    deleteRowRuns_(o.routine, rows);
  }
  [['health', o.health, 3], ['check', o.check, 3], ['indiv', o.indiv, 1]].forEach(function (d) {
    var last = d[1].getLastRow();
    if (last < 2) return;
    var rows = [];
    d[1].getRange(2, d[2], last - 1, 1).getValues().forEach(function (r, i) { if (String(r[0]) === id) rows.push(i + 2); });
    out[d[0]] += rows.length;
    deleteRowRuns_(d[1], rows);
  });
  return out;
}

function findStudentFresh_(ctx, id) {
  var st = null;
  readRosterRaw_(true).forEach(function (s) { if (s.id === id) st = s; });
  if (!st) fail_('ไม่พบนักเรียนคนนี้ (อาจถูกลบไปแล้ว) กรุณาโหลดหน้าใหม่', 'BAD');
  if (!roomAllowed_(ctx, st.room)) fail_('ไม่มีสิทธิ์ลบนักเรียนนอกห้องที่ได้รับสิทธิ์', 'FORBIDDEN');
  return st;
}

function previewDeleteStudentFor_(ctx, id) {
  id = assertId_(ctx, id);
  var st = findStudentFresh_(ctx, id);
  return { id: st.id, name: st.name, room: st.room, counts: countStudentData_(id) };
}

// ลบนักเรียน 1 คน: ข้อมูลนม/แปรงฟัน/น้ำหนัก-ส่วนสูง/ตรวจสุขภาพ/ประวัติรายบุคคล แล้วจึงลบจากรายชื่อ
// ลำดับนี้ตั้งใจ: ถ้าหยุดกลางทาง นักเรียนยังอยู่ในรายชื่อ กดลบซ้ำได้ (ไม่เหลือข้อมูลค้างที่ไม่มีเจ้าของ)
// ต้องส่ง confirmFlag === true (หน้าเว็บถามผู้ใช้ก่อนเสมอ) — กู้คืนได้เฉพาะจาก "ประวัติเวอร์ชัน" ของ Google Sheet
function deleteStudentFor_(ctx, id, confirmFlag) {
  if (confirmFlag !== true) fail_('ต้องยืนยันการลบก่อน', 'BAD');
  id = assertId_(ctx, id);
  return withLock_(function () {
    var st = findStudentFresh_(ctx, id);
    var removed = removeStudentData_(id);
    var list = readRosterRaw_(true).filter(function (s) { return s.id !== id; });
    writeRoster_(ensureSetup_().roster, list);
    var res = { success: true, deleted: removed, name: st.name, room: st.room,
                message: 'ลบ "' + st.name + '" ออกจากระบบแล้ว (พร้อมข้อมูลที่เกี่ยวข้องในชีต)',
                students: getRoster_(true), revs: roomRevs_(readRosterRaw_(true), getRooms_(true)) };
    return scopeResult_(ctx, res);
  }, ['roster', 'data']);
}
function importStudentsFor_(ctx, items, defaultRoom, year, month) {
  if (!Array.isArray(items) || items.length > 2000) fail_('ข้อมูลนำเข้าไม่ถูกต้อง (สูงสุด 2,000 แถว)', 'BAD');
  defaultRoom = String(defaultRoom == null ? '' : defaultRoom).trim();
  if (!ctx.allRooms && !roomAllowed_(ctx, defaultRoom)) fail_('ไม่มีสิทธิ์นำเข้าข้อมูลห้องนี้', 'FORBIDDEN');
  if (year || month) { year = intIn_(year, 2400, 2800, 'ปีการศึกษา'); month = intIn_(month, 1, 12, 'เดือน'); }
  items = items.map(function (it) {
    it = plainObj_(it);
    var o = cleanStudent_(it);
    o.room = clip_(it.room, 40).trim();
    var w = parseFloat(it.w), h = parseFloat(it.h);
    if (!isNaN(w)) o.w = w;
    if (!isNaN(h)) o.h = h;
    return o;
  });
  return scopeResult_(ctx, importStudents_(items, defaultRoom, year, month, ctx));
}

// ---- เฉพาะ Admin ----
function saveSettingsFor_(ctx, obj) {
  obj = plainObj_(obj);
  var o = {};
  Object.keys(obj).forEach(function (k) { o[k] = typeof obj[k] === 'number' ? obj[k] : clip_(obj[k], 200); });
  if (o.academicYear !== undefined) intIn_(o.academicYear, 2400, 2700, 'ปีการศึกษา');
  if (o.startMonth !== undefined) intIn_(o.startMonth, 1, 12, 'เดือนเริ่มต้น');
  if (o.currentMonth !== undefined) intIn_(o.currentMonth, 1, 12, 'เดือนปัจจุบัน');
  return saveSettings_(o);
}
function saveRoomsFor_(ctx, rooms) {
  if (!Array.isArray(rooms) || rooms.length > 80) fail_('ข้อมูลห้องเรียนไม่ถูกต้อง', 'BAD');
  return saveRooms_(rooms.map(function (r) {
    r = plainObj_(r);
    return { old: clip_(r.old, 40), name: clip_(r.name, 40), level: clip_(r.level, 10), suffix: clip_(r.suffix, 12),
             teacher: clip_(r.teacher, 100), pos: clip_(r.pos, 120) };
  }));
}
function promoteYearFor_(ctx, fromYear, removeGraduates) {
  return promoteYear_(intIn_(fromYear, 2400, 2800, 'ปีการศึกษา'), removeGraduates === true);
}
function saveHolidaysFor_(ctx, rows) {
  if (!Array.isArray(rows) || rows.length > 500) fail_('ข้อมูลวันหยุดไม่ถูกต้อง', 'BAD');
  return saveHolidays_(rows.map(function (r) {
    r = plainObj_(r);
    return { date: clip_(r.date, 10), name: clip_(r.name, 120), open: r.open === true };
  }));
}

// ---- จัดการผู้ใช้ (เฉพาะ Admin) ----
function listUsersFor_(ctx) { return usersPayload_(); }

function saveUserFor_(ctx, o) {
  o = plainObj_(o);
  var username = String(o.username == null ? '' : o.username).trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) fail_('Username ต้องเป็นตัวอักษรอังกฤษพิมพ์เล็ก ตัวเลข จุด ขีดกลาง หรือขีดล่าง ยาว 3-30 ตัว', 'BAD');
  var name = clip_(o.name, 80).trim();
  if (!name) fail_('กรุณากรอกชื่อที่แสดง', 'BAD');
  var role = String(o.role == null ? '' : o.role);
  if (role !== ROLE_TEACHER && role !== ROLE_DIRECTOR) fail_('บทบาทไม่ถูกต้อง', 'BAD');
  var rooms = [];
  if (role === ROLE_TEACHER) {
    var valid = {};
    getRooms_(true).forEach(function (r) { valid[r.name] = 1; });
    strArr_(o.rooms == null ? [] : o.rooms, 'ห้องเรียน', 100).forEach(function (n) {
      if (!valid[n]) fail_('ไม่พบห้อง "' + n + '"', 'BAD');
      if (rooms.indexOf(n) < 0) rooms.push(n);
    });
    if (!rooms.length) fail_('ครูต้องได้รับสิทธิ์อย่างน้อย 1 ห้อง', 'BAD');
  }
  var pin = String(o.pin == null ? '' : o.pin);
  return withLock_(function () {
    var sh = ensureUsers_(SpreadsheetApp.getActiveSpreadsheet());
    var u = findUser_(username, true), t = stamp_();
    if (o.isNew === true) {
      if (u) fail_('Username นี้มีอยู่แล้ว', 'BAD');
      var pp = pinProblem_(pin);
      if (pp) fail_(pp, 'BAD');
      var salt = randHex_().substring(0, 32);
      var at = Math.max(sh.getLastRow(), 1) + 1;
      ensureRows_(sh, at);
      // รอบเซสชันเริ่มจากเวลาปัจจุบัน (ไม่ใช่ 0) เพื่อให้เซสชันของบัญชีเก่าชื่อเดียวกันที่ถูกลบไปแล้ว ใช้กับบัญชีใหม่ไม่ได้
      sh.getRange(at, 1, 1, USER_HEADERS.length).setValues([[username, name, role, JSON.stringify(rooms), hashPin_(pin, salt, username),
        salt, 'active', 0, '', Math.floor(now_() / 1000), t, t]]);
      return { success: true, message: 'สร้างบัญชี "' + username + '" เรียบร้อยแล้ว', users: usersPayload_(true).users };
    }
    if (!u) fail_('ไม่พบผู้ใช้นี้', 'BAD');
    var ver = u.ver, status = u.status, fails = u.fails, hash = u.hash, salt2 = u.salt, pinChanged = false;
    if (role !== u.role) ver++;
    if (pin !== '') {
      var pp2 = pinProblem_(pin);
      if (pp2) fail_(pp2, 'BAD');
      salt2 = randHex_().substring(0, 32);
      hash = hashPin_(pin, salt2, username);
      ver++; fails = 0; pinChanged = true;
      if (status === 'locked') status = 'active';
    }
    sh.getRange(u.row, UCOL.name, 1, 3).setValues([[name, role, JSON.stringify(rooms)]]);
    sh.getRange(u.row, UCOL.hash, 1, 6).setValues([[hash, salt2, status, fails, u.lastLogin, ver]]);
    sh.getRange(u.row, UCOL.updated).setValue(t);
    return { success: true, message: 'บันทึกบัญชี "' + username + '" แล้ว' + (pinChanged ? ' (ตั้ง PIN ใหม่ — ผู้ใช้ที่ล็อกอินอยู่จะถูกออกจากระบบ)' : ''), users: usersPayload_(true).users };
  }, ['users']);
}

function setUserStatusFor_(ctx, username, status) {
  username = String(username == null ? '' : username).trim().toLowerCase();
  if (status !== 'active' && status !== 'disabled') fail_('สถานะไม่ถูกต้อง', 'BAD');
  return withLock_(function () {
    var sh = ensureUsers_(SpreadsheetApp.getActiveSpreadsheet());
    var u = findUser_(username, true);
    if (!u) fail_('ไม่พบผู้ใช้นี้', 'BAD');
    var t = stamp_();
    if (status === 'disabled') {
      sh.getRange(u.row, UCOL.status).setValue('disabled');
      sh.getRange(u.row, UCOL.ver).setValue(u.ver + 1);
    } else {
      sh.getRange(u.row, UCOL.status, 1, 2).setValues([['active', 0]]);
    }
    sh.getRange(u.row, UCOL.updated).setValue(t);
    return { success: true, message: status === 'disabled' ? 'ปิดใช้งานบัญชี "' + username + '" แล้ว' : 'เปิดใช้งาน/ปลดล็อกบัญชี "' + username + '" แล้ว', users: usersPayload_(true).users };
  }, ['users']);
}

// ลบบัญชีผู้ใช้ถาวร (เฉพาะ Admin): ต้องพิมพ์ Username ยืนยัน (ตรวจซ้ำที่เซิร์ฟเวอร์) — ลบแถวออกจากชีต "ผู้ใช้งาน" จริง
// เซสชันที่ล็อกอินอยู่ของบัญชีนี้ใช้ต่อไม่ได้ทันที เพราะ authenticate_ หาผู้ใช้ไม่พบ (แคชผู้ใช้ถูกล้างใน withLock_)
// ไม่กระทบข้อมูลนักเรียน/ข้อมูลบันทึกใด ๆ (ข้อมูลเหล่านั้นไม่ผูกกับบัญชีผู้ใช้)
function deleteUserFor_(ctx, username, confirmName) {
  username = String(username == null ? '' : username).trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) fail_('Username ไม่ถูกต้อง', 'BAD');
  if (String(confirmName == null ? '' : confirmName).trim().toLowerCase() !== username) {
    fail_('ชื่อที่พิมพ์ยืนยันไม่ตรงกับ Username ที่จะลบ จึงไม่ได้ลบ', 'BAD');
  }
  return withLock_(function () {
    var sh = ensureUsers_(SpreadsheetApp.getActiveSpreadsheet());
    var u = findUser_(username, true);
    if (!u) fail_('ไม่พบผู้ใช้นี้ (อาจถูกลบไปแล้ว)', 'BAD');
    deleteRowRuns_(sh, [u.row]);
    return { success: true, message: 'ลบบัญชี "' + username + '" ออกจากระบบแล้ว (ผู้ใช้ที่ล็อกอินอยู่จะถูกออกจากระบบทันที)', users: usersPayload_(true).users };
  }, ['users']);
}

/* ---------- ประตูเดียวที่หน้าเว็บเรียกได้ ---------- */

// ตารางสิทธิ์: ชื่อคำสั่ง -> บทบาทที่ใช้ได้ + ฟังก์ชันที่รัน (ctx = ผลตรวจสิทธิ์ฝั่งเซิร์ฟเวอร์)
var RPC = {
  ping:            { roles: R_ALL,   run: function (c, a) { return { alive: true }; } },
  getBootstrap:    { roles: R_ALL,   run: function (c, a) { return getBootstrapFor_(c, a[0] === true); } },
  getPeriod:       { roles: R_ALL,   run: function (c, a) { return getPeriodFor_(c, a[0], a[1]); } },
  getGrowth:       { roles: R_ALL,   run: function (c, a) { return getGrowthFor_(c, a[0], a[1]); } },
  getIndividual:   { roles: R_ALL,   run: function (c, a) { return getIndividualFor_(c, a[0]); } },
  saveRoutine:     { roles: R_WRITE, run: function (c, a) { return saveRoutineFor_(c, a[0], a[1], a[2], a[3], a[4], a[5]); } },
  saveCheck:       { roles: R_WRITE, run: function (c, a) { return saveCheckFor_(c, a[0], a[1], a[2], a[3], a[4]); } },
  saveHealth:      { roles: R_WRITE, run: function (c, a) { return saveHealthFor_(c, a[0], a[1], a[2], a[3], a[4]); } },
  saveIndividual:  { roles: R_WRITE, run: function (c, a) { return saveIndividualFor_(c, a[0], a[1]); } },
  clearIndividual: { roles: R_WRITE, run: function (c, a) { return clearIndividualFor_(c, a[0]); } },
  saveRoster:      { roles: R_WRITE, run: function (c, a) { return saveRosterFor_(c, a[0], a[1], a[2]); } },
  previewDeleteStudent: { roles: R_WRITE, run: function (c, a) { return previewDeleteStudentFor_(c, a[0]); } },
  deleteStudent:   { roles: R_WRITE, run: function (c, a) { return deleteStudentFor_(c, a[0], a[1]); } },
  importStudents:  { roles: R_WRITE, run: function (c, a) { return importStudentsFor_(c, a[0], a[1], a[2], a[3]); } },
  saveSettings:    { roles: R_ADMIN, run: function (c, a) { return saveSettingsFor_(c, a[0]); } },
  saveRooms:       { roles: R_ADMIN, run: function (c, a) { return saveRoomsFor_(c, a[0]); } },
  promoteYear:     { roles: R_ADMIN, run: function (c, a) { return promoteYearFor_(c, a[0], a[1]); } },
  saveHolidays:    { roles: R_ADMIN, run: function (c, a) { return saveHolidaysFor_(c, a[0]); } },
  listUsers:       { roles: R_ADMIN, run: function (c, a) { return listUsersFor_(c); } },
  saveUser:        { roles: R_ADMIN, run: function (c, a) { return saveUserFor_(c, a[0]); } },
  setUserStatus:   { roles: R_ADMIN, run: function (c, a) { return setUserStatusFor_(c, a[0], a[1]); } },
  deleteUser:      { roles: R_ADMIN, run: function (c, a) { return deleteUserFor_(c, a[0], a[1]); } }
};

function errResp_(e) {
  if (e && e.appCode) return { ok: false, code: e.appCode, message: e.message };
  try { console.error(String(e && e.stack ? e.stack : e)); } catch (x) {}
  var msg = String(e && e.message ? e.message : e);
  if (/lock/i.test(msg)) return { ok: false, code: 'BUSY', message: 'ระบบกำลังบันทึกข้อมูลอื่นอยู่ กรุณาลองใหม่อีกครั้ง' };
  return { ok: false, code: 'ERR', message: 'เกิดข้อผิดพลาดในระบบ: ' + msg.substring(0, 150) };
}

// ฟังก์ชันเดียวที่หน้าเว็บเรียกได้ (google.script.run.rpc) — ไม่มีคำสั่งใดทำงานได้ถ้าไม่ผ่านการตรวจสิทธิ์ด้านบน
function rpc(token, fn, args) {
  try {
    fn = String(fn == null ? '' : fn);
    args = Array.isArray(args) ? args : [];
    if (fn === 'hello') return ok_(helloInfo_());
    if (fn === 'login') return loginRpc_(args[0], args[1]);
    if (fn === 'adminEnter') return adminEnterRpc_();
    if (fn === 'logout') { killSession_(token); return ok_({ bye: true }); }
    if (!Object.prototype.hasOwnProperty.call(RPC, fn)) fail_('ไม่พบคำสั่งที่เรียก', 'BAD');
    var def = RPC[fn];
    var ctx = authenticate_(token);
    if (def.roles.indexOf(ctx.role) < 0) fail_('บัญชีนี้ไม่มีสิทธิ์ทำรายการนี้', 'FORBIDDEN');
    return ok_(def.run(ctx, args));
  } catch (e) {
    return errResp_(e);
  }
}

/* ------------------------------------------------------------------ */
/* ------------------------------------------------------------------ */
/*  ฟังก์ชันงาน (เรียกได้ผ่าน rpc() ที่ตรวจสิทธิ์แล้วเท่านั้น)                 */
/* ------------------------------------------------------------------ */

function getBootstrap_() {
  ensureSetup_();
  return {
    settings: getSettings_(),
    rooms: getRooms_(),
    students: getRoster_(),
    holidays: getHolidays_(),
    criteria: getCriteria_(),
    years: getYears_()
  };
}

// ข้อมูลของเดือนหนึ่ง (ผ่านแคชกลุ่ม data+meta อายุสั้น — ถูกล้างทุกครั้งที่มีการบันทึกข้อมูล)
function getPeriod_(year, month) {
  year = Number(year); month = Number(month);
  return cached_(['data', 'meta'], 'period:' + year + ':' + month, CACHE_TTL.period, function () {
    var r = readRoutine_(year, month);
    return { milk: r.milk, brush: r.brush, health: readHealth_(year, month), ref: heightRef_(year, month), check: readCheck_(year, month) };
  });
}

/* ---------- เขียนเฉพาะแถวที่เปลี่ยน (แทนการล้างทั้งชีตแล้วเขียนใหม่) ---------- */
// ทุกฟังก์ชันที่ใช้ตัวช่วยนี้ต้องอ่านข้อมูลและเขียน "ภายในล็อกเดียวกัน" (withLock_) เพื่อกันข้อมูลทับกัน
// data = ค่าทั้งชีตตั้งแต่แถว 2 ที่อ่านไว้แล้ว  mutFrom = คอลัมน์แรก (นับจาก 1) ที่ค่าอาจเปลี่ยน
// updates = { เลขแถว: [ค่าทั้งแถว] }  appends = [[ค่าทั้งแถว], ...]  removes = [เลขแถว, ...]
// - updates: รวมแถวที่อยู่ใกล้กัน (ห่างไม่เกิน 25 แถว) เป็นก้อนเดียว เขียนครั้งเดียวต่อก้อน เฉพาะคอลัมน์ที่เปลี่ยนได้
// - appends: ต่อท้ายชีตครั้งเดียว   - removes: ลบแถวจริง (ไม่ใช่แค่ล้างค่า) จากล่างขึ้นบน
function applyRowOps_(sh, width, mutFrom, data, updates, appends, removes) {
  var rows = Object.keys(updates).map(Number).sort(function (a, b) { return a - b; });
  var i = 0;
  while (i < rows.length) {
    var j = i;
    while (j + 1 < rows.length && rows[j + 1] - rows[j] <= 25) j++;
    var lo = rows[i], hi = rows[j], block = [];
    for (var r = lo; r <= hi; r++) block.push((updates[r] || data[r - 2]).slice(mutFrom - 1));
    sh.getRange(lo, mutFrom, hi - lo + 1, width - mutFrom + 1).setValues(block);
    i = j + 1;
  }
  if (appends.length) {
    var at = sh.getLastRow() + 1;
    ensureRows_(sh, at + appends.length);
    sh.getRange(at, 1, appends.length, width).setValues(appends);
  }
  if (removes.length) deleteRowRuns_(sh, removes);
}

// ลบแถวตามเลขแถวที่ระบุ (รวมแถวติดกันเป็นก้อนเดียว) ลบจากล่างขึ้นบนเพื่อไม่ให้เลขแถวด้านบนเลื่อน
function deleteRowRuns_(sh, rows) {
  rows = rows.slice().sort(function (a, b) { return b - a; }).filter(function (r, k, a) { return k === 0 || r !== a[k - 1]; });
  if (!rows.length) return;
  ensureRows_(sh, sh.getLastRow() + 2);   // Sheets ไม่ยอมลบ "ทุกแถวที่ไม่ได้ freeze" จึงเผื่อแถวว่างไว้เสมอ
  var i = 0;
  while (i < rows.length) {
    var j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] - 1) j++;
    sh.deleteRows(rows[j], j - i + 1);
    i = j + 1;
  }
}

function parseDays_(v) {
  var seen = {}, out = [];
  String(v == null ? '' : v).split(',').forEach(function (x) {
    x = String(x).trim();
    if (x === '') return;
    var n = Number(x);
    if (isFinite(n) && n >= 1 && n <= 31 && Math.floor(n) === n && !seen[n]) { seen[n] = 1; out.push(n); }
  });
  return out.sort(function (a, b) { return a - b; });
}
function daysList_(arr) {
  var seen = {}, out = [];
  (Array.isArray(arr) ? arr : []).forEach(function (d) {
    d = Number(d);
    if (isFinite(d) && d >= 1 && d <= 31 && Math.floor(d) === d && !seen[d]) { seen[d] = 1; out.push(d); }
  });
  return out.sort(function (a, b) { return a - b; });
}

// บันทึกนม/แปรงฟัน 1 เดือน ภายในล็อก: อ่านข้อมูลล่าสุดจากชีต แล้วเขียนเฉพาะแถวที่เปลี่ยน
//  mode 'diff' (หน้าเว็บใหม่): grid = { id: {a:[วันที่เพิ่ม], r:[วันที่เอาออก]} } — ปรับเฉพาะวันที่ที่ผู้ใช้เปลี่ยน
//     ผสานกับข้อมูลล่าสุดในชีต จึงไม่ลบข้อมูลที่คนอื่น/อุปกรณ์อื่นเพิ่งบันทึก และบันทึกซ้ำได้ปลอดภัย
//  mode อื่น (เดิม): grid = { id: [วันที่ทั้งหมด] } แทนที่ข้อมูลของ ids ทั้งชุด (คงไว้เพื่อความเข้ากันได้)
// ids = รหัสนักเรียนของห้องที่กำลังบันทึก (ไม่แตะข้อมูลห้องอื่น) — คืน merged = วันที่ล่าสุดของ ids ทุกคนหลังบันทึก
function saveRoutine_(year, month, type, grid, ids, mode) {
  year = Number(year); month = Number(month);
  if (type !== 'milk' && type !== 'brush') return { success: false, message: 'ประเภทข้อมูลไม่ถูกต้อง' };
  var diff = (mode === 'diff');
  return withLock_(function () {
    var sh = ensureSetup_().routine, last = sh.getLastRow();
    var data = last > 1 ? sh.getRange(2, 1, last - 1, 6).getValues() : [];
    var scope = ids ? idSet_(ids) : null, rowOf = {}, removes = [];
    data.forEach(function (r, i) {
      if (Number(r[0]) !== year || Number(r[1]) !== month || String(r[2]) !== type) return;
      var id = String(r[3]);
      if (id === '' || (scope && !scope[id])) return;
      if (rowOf[id] !== undefined) { removes.push(i + 2); return; }   // แถวซ้ำของคนเดียวกัน: เก็บแถวแรก ลบที่เหลือ
      rowOf[id] = i;
    });
    var want = {};   // id -> วันที่ที่ควรมีหลังบันทึก (เรียงแล้ว)
    grid = grid || {};
    if (!diff) Object.keys(rowOf).forEach(function (id) { want[id] = []; });
    Object.keys(grid).forEach(function (id) {
      if (!diff) { want[id] = daysList_(grid[id]); return; }
      var m = {}, g = grid[id] || {};
      (rowOf[id] !== undefined ? parseDays_(data[rowOf[id]][4]) : []).forEach(function (d) { m[d] = 1; });
      daysList_(g.r).forEach(function (d) { delete m[d]; });
      daysList_(g.a).forEach(function (d) { m[d] = 1; });
      want[id] = Object.keys(m).map(Number).sort(function (a, b) { return a - b; });
    });
    var updates = {}, appends = [];
    Object.keys(want).forEach(function (id) {
      var days = want[id], i = rowOf[id], str = days.join(',');
      if (i === undefined) { if (days.length) appends.push([year, month, type, id, str, days.length]); return; }
      if (!days.length) { removes.push(i + 2); return; }
      var cur = data[i];
      if (String(cur[4]) !== str || Number(cur[5]) !== days.length) { var nr = cur.slice(); nr[4] = str; nr[5] = days.length; updates[i + 2] = nr; }
    });
    applyRowOps_(sh, 6, 5, data, updates, appends, removes);
    var merged = {};
    (ids || Object.keys(want)).forEach(function (id) {
      var d = want[id] !== undefined ? want[id] : (rowOf[id] !== undefined ? parseDays_(data[rowOf[id]][4]) : []);
      if (d.length) merged[id] = d;
    });
    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว', merged: merged };
  }, ['data']);
}

// ลำดับเดือนตามปีการศึกษา (ใช้เรียงเวลาข้ามปี)
function monthOrd_(y, m, sm) { return Number(y) * 12 + ((Number(m) - sm + 12) % 12); }

// คำนวณค่าอ้างอิงจากแถวที่อ่านไว้แล้ว (ใช้ซ้ำได้ทั้งตอนอ่านและตอนตรวจก่อนบันทึกภายในล็อก)
function heightRefFrom_(rows, year, month, sm) {
  var cur = monthOrd_(year, month, sm), ref = {};
  rows.forEach(function (r) {
    var id = String(r[2]);
    if (!id) return;
    var o = monthOrd_(r[0], r[1], sm);
    if (o === cur) return;
    var e = ref[id] || (ref[id] = { prevH: null, nextH: null, prevW: null, nextW: null });
    [['H', r[4]], ['W', r[3]]].forEach(function (p) {
      var v = p[1] === '' ? NaN : Number(p[1]);
      if (isNaN(v)) return;
      var item = { y: Number(r[0]), m: Number(r[1]), v: v, ord: o };
      if (o < cur) { if (!e['prev' + p[0]] || o > e['prev' + p[0]].ord) e['prev' + p[0]] = item; }
      else { if (!e['next' + p[0]] || o < e['next' + p[0]].ord) e['next' + p[0]] = item; }
    });
  });
  return ref;
}

// ค่าน้ำหนัก/ส่วนสูงของเดือนก่อนหน้า (ล่าสุดก่อนเดือนนี้) และเดือนถัดไป (ใกล้สุดหลังเดือนนี้) ของนักเรียนแต่ละคน
function heightRef_(year, month) {
  var sm = Number(getSettings_().startMonth) || 5;
  var sh = ensureSetup_().health, last = sh.getLastRow();
  if (last < 2) return {};
  return heightRefFrom_(sh.getRange(2, 1, last - 1, 5).getValues(), year, month, sm);
}

// ตรวจสุขภาพรายสัปดาห์: 1 แถวต่อคนต่อเดือน, ผลเป็นสตริง 20 ตัวอักษร (g=ดี, x=ปรับปรุง, -=ว่าง)
function readCheck_(year, month) {
  var sh = ensureSetup_().check, out = {}, last = sh.getLastRow();
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) {
    if (Number(r[0]) === Number(year) && Number(r[1]) === Number(month) && String(r[2]) !== '') out[String(r[2])] = fixCheck_(r[3]);
  });
  return out;
}
function fixCheck_(v) {
  var t = String(v == null ? '' : v).replace(/[^gx\-]/g, '');
  while (t.length < 20) t += '-';
  return t.substring(0, 20);
}
// ตรวจสุขภาพรายสัปดาห์: บันทึกเฉพาะแถวที่เปลี่ยน ภายในล็อก
//  mode 'diff': map = { id: { ตำแหน่ง0-19: 'g'|'x'|'-' } } ปรับเฉพาะช่องที่ผู้ใช้เปลี่ยน ผสานกับค่าล่าสุดในชีต
//  mode อื่น (เดิม): map = { id: สตริง 20 ตัว } แทนที่ข้อมูลของ ids ทั้งชุด
function saveCheck_(year, month, map, ids, mode) {
  year = Number(year); month = Number(month);
  var diff = (mode === 'diff');
  return withLock_(function () {
    var sh = ensureSetup_().check, last = sh.getLastRow();
    var data = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
    var scope = ids ? idSet_(ids) : null, rowOf = {}, removes = [];
    data.forEach(function (r, i) {
      if (Number(r[0]) !== year || Number(r[1]) !== month) return;
      var id = String(r[2]);
      if (id === '' || (scope && !scope[id])) return;
      if (rowOf[id] !== undefined) { removes.push(i + 2); return; }
      rowOf[id] = i;
    });
    var want = {};   // id -> สตริง 20 ตัวที่ควรมีหลังบันทึก
    map = map || {};
    if (!diff) Object.keys(rowOf).forEach(function (id) { want[id] = fixCheck_(''); });
    Object.keys(map).forEach(function (id) {
      if (!diff) { want[id] = fixCheck_(map[id]); return; }
      var a = fixCheck_(rowOf[id] !== undefined ? data[rowOf[id]][3] : '').split(''), df = map[id] || {};
      Object.keys(df).forEach(function (k) {
        var p = Number(k), c = String(df[k]);
        if (p >= 0 && p < 20 && Math.floor(p) === p && (c === 'g' || c === 'x' || c === '-')) a[p] = c;
      });
      want[id] = a.join('');
    });
    var updates = {}, appends = [];
    Object.keys(want).forEach(function (id) {
      var t = want[id], i = rowOf[id], empty = /^-+$/.test(t);
      if (i === undefined) { if (!empty) appends.push([year, month, id, t]); return; }
      if (empty) { removes.push(i + 2); return; }
      var cur = data[i];
      if (String(cur[3]) !== t) { var nr = cur.slice(); nr[3] = t; updates[i + 2] = nr; }
    });
    applyRowOps_(sh, 4, 4, data, updates, appends, removes);
    var merged = {};
    (ids || Object.keys(want)).forEach(function (id) {
      var t = want[id] !== undefined ? want[id] : (rowOf[id] !== undefined ? fixCheck_(data[rowOf[id]][3]) : '');
      if (t && !/^-+$/.test(t)) merged[id] = t;
    });
    return { success: true, message: 'บันทึกการตรวจสุขภาพเรียบร้อยแล้ว', merged: merged };
  }, ['data']);
}

function hnum_(v) { var n = parseFloat(v); return isNaN(n) ? '' : n; }

// วางแผนการเขียนน้ำหนัก/ส่วนสูงของเดือนหนึ่ง จากแถวที่อ่านไว้แล้ว (ใช้ร่วมกันระหว่างบันทึกปกติและนำเข้ารายชื่อ)
// patches = { id: {w?, h?} } เฉพาะช่องที่ระบุเท่านั้นที่เปลี่ยน ('' = ล้างช่อง) ช่องที่ไม่ระบุคงค่าเดิมในชีต
// replaceIds (ถ้ามี) = โหมดเดิม: นักเรียนใน replaceIds ที่ไม่มีใน patches จะถูกลบแถว
function planHealth_(data, year, month, patches, replaceIds) {
  var rowOf = {}, removes = [];
  data.forEach(function (r, i) {
    if (Number(r[0]) !== year || Number(r[1]) !== month) return;
    var id = String(r[2]);
    if (id === '') return;
    if (rowOf[id] !== undefined) { if (!replaceIds || replaceIds[id]) removes.push(i + 2); return; }
    rowOf[id] = i;
  });
  var want = {};
  if (replaceIds) Object.keys(rowOf).forEach(function (id) { if (replaceIds[id]) want[id] = { w: '', h: '' }; });
  Object.keys(patches).forEach(function (id) {
    var p = patches[id], i = rowOf[id], cur = i !== undefined ? data[i] : null;
    var v = want[id] || { w: cur ? cur[3] : '', h: cur ? cur[4] : '' };
    if (Object.prototype.hasOwnProperty.call(p, 'w')) v.w = hnum_(p.w);
    if (Object.prototype.hasOwnProperty.call(p, 'h')) v.h = hnum_(p.h);
    want[id] = v;
  });
  var updates = {}, appends = [];
  Object.keys(want).forEach(function (id) {
    var v = want[id], i = rowOf[id], empty = (v.w === '' && v.h === '');
    if (i === undefined) { if (!empty) appends.push([year, month, id, v.w, v.h]); return; }
    if (empty) { removes.push(i + 2); return; }
    var cur = data[i];
    if (String(cur[3]) !== String(v.w) || String(cur[4]) !== String(v.h)) { var nr = cur.slice(); nr[3] = v.w; nr[4] = v.h; updates[i + 2] = nr; }
  });
  return { rowOf: rowOf, want: want, updates: updates, appends: appends, removes: removes };
}

// บันทึกน้ำหนัก-ส่วนสูงของเดือนหนึ่ง ภายในล็อก (เขียนเฉพาะแถวที่เปลี่ยน)
//  mode 'diff': rows = [{id, w?, h?}] เฉพาะช่องที่ผู้ใช้แก้ ผสานกับค่าล่าสุดในชีต
//  mode อื่น (เดิม): rows = [{id, w, h}] ของ ids ทั้งชุด แทนที่ทั้งหมด
function saveHealth_(year, month, rows, ids, mode) {
  year = Number(year); month = Number(month);
  var diff = (mode === 'diff');
  return withLock_(function () {
    var sh = ensureSetup_().health, last = sh.getLastRow();
    var data = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
    var has = function (o, k) { return Object.prototype.hasOwnProperty.call(o, k); };
    var patches = {};
    (rows || []).forEach(function (r) {
      var p = patches[String(r.id)] || (patches[String(r.id)] = {});
      if (!diff || has(r, 'w')) p.w = r.w;
      if (!diff || has(r, 'h')) p.h = r.h;
    });
    // ตรวจส่วนสูงซ้ำที่เซิร์ฟเวอร์ (ข้อมูลล่าสุดภายในล็อก): ส่วนสูงลดลงจากเดือนก่อนหน้า หรือสูงกว่าเดือนถัดไปไม่ได้
    var sm = Number(getSettings_(true).startMonth) || 5;
    var ref = heightRefFrom_(data, year, month, sm), bad = [], names = {};
    readRosterRaw_().forEach(function (st) { names[st.id] = st.name; });
    Object.keys(patches).forEach(function (id) {
      if (!has(patches[id], 'h')) return;
      var h = parseFloat(patches[id].h), e = ref[id];
      if (isNaN(h)) return;
      if (h < 40 || h > 200) bad.push((names[id] || id) + ' (ส่วนสูงต้อง 40-200 ซม.)');
      else if (e && e.prevH && h < e.prevH.v) bad.push((names[id] || id) + ' (ต่ำกว่าเดือนก่อนหน้า ' + e.prevH.v + ' ซม.)');
      else if (e && e.nextH && h > e.nextH.v) bad.push((names[id] || id) + ' (สูงกว่าเดือนถัดไป ' + e.nextH.v + ' ซม.)');
    });
    if (bad.length) return { success: false, message: 'ส่วนสูงไม่สมเหตุสมผล: ' + bad.slice(0, 5).join(', ') + (bad.length > 5 ? ' และอีก ' + (bad.length - 5) + ' คน' : '') };
    var plan = planHealth_(data, year, month, patches, (!diff && ids) ? idSet_(ids) : null);
    applyRowOps_(sh, 5, 4, data, plan.updates, plan.appends, plan.removes);
    var merged = {};
    (ids || Object.keys(plan.want)).forEach(function (id) {
      var v = plan.want[id];
      if (!v && plan.rowOf[id] !== undefined) v = { w: data[plan.rowOf[id]][3], h: data[plan.rowOf[id]][4] };
      if (v && !(v.w === '' && v.h === '')) merged[id] = { w: v.w === '' ? '' : Number(v.w), h: v.h === '' ? '' : Number(v.h) };
    });
    return { success: true, message: 'บันทึกน้ำหนัก-ส่วนสูงเรียบร้อยแล้ว', merged: merged };
  }, ['data']);
}

function getGrowth_(year, id) {
  year = Number(year);
  var sh = ensureSetup_().health;
  var last = sh.getLastRow();
  var out = [];
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 5).getValues().forEach(function (r) {
    if (Number(r[0]) === year && String(r[2]) === String(id)) {
      out.push({ month: Number(r[1]), w: r[3] === '' ? null : Number(r[3]), h: r[4] === '' ? null : Number(r[4]) });
    }
  });
  return out;
}

function getIndividual_(id) {
  var sh = ensureSetup_().indiv;
  var last = sh.getLastRow();
  if (last < 2) return {};
  var v = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) {
      try { return JSON.parse(v[i][1]); } catch (e) { return {}; }
    }
  }
  return {};
}

function saveIndividual_(id, data) {
  id = String(id == null ? '' : id).trim();
  if (!id || id === 'undefined' || id === 'null') return { success: false, message: 'ไม่ได้เลือกนักเรียน' };
  return withLock_(function () {
    var sh = ensureSetup_().indiv;
    var last = sh.getLastRow();
    var json = JSON.stringify(data || {});
    var stamp = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
    var row = -1;
    if (last > 1) {
      var v = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < v.length; i++) { if (String(v[i][0]) === String(id)) { row = i + 2; break; } }
    }
    if (row > 0) sh.getRange(row, 2, 1, 2).setValues([[json, stamp]]);
    else { ensureRows_(sh, last + 1); sh.getRange(last + 1, 1, 1, 3).setValues([[String(id), json, stamp]]); }
    return { success: true, message: 'บันทึกประวัติสุขภาพเรียบร้อยแล้ว' };
  });
}

// ล้างแบบประวัติสุขภาพรายบุคคลของนักเรียนคนนี้ (ไม่แตะข้อมูลนม แปรงฟัน น้ำหนัก-ส่วนสูง)
function clearIndividual_(id) {
  id = String(id == null ? '' : id).trim();
  if (!id) return { success: false, message: 'ไม่ได้เลือกนักเรียน' };
  return withLock_(function () {
    var sh = ensureSetup_().indiv;
    var last = sh.getLastRow();
    if (last > 1) {
      var v = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = v.length - 1; i >= 0; i--) { if (String(v[i][0]) === id) sh.deleteRow(i + 2); }
    }
    return { success: true, message: 'ล้างแบบประวัติสุขภาพของนักเรียนคนนี้แล้ว' };
  });
}

function saveSettings_(obj) {
  return withLock_(function () {
    var sh = ensureSetup_().settings;
    var last = sh.getLastRow();
    var vals = last >= 3 ? sh.getRange(3, 1, last - 2, 3).getValues() : [];
    var out = vals.map(function (r) {
      var k = String(r[2]);
      return [(obj && Object.prototype.hasOwnProperty.call(obj, k)) ? obj[k] : r[1]];
    });
    if (out.length) sh.getRange(3, 2, out.length, 1).setValues(out);
    return { success: true, message: 'บันทึกการตั้งค่าเรียบร้อยแล้ว', settings: getSettings_(true) };
  }, ['meta']);
}

function saveRoster_(rows, room, baseRev) {
  return withLock_(function () {
    var o = ensureSetup_();
    var rooms = getRooms_(true);
    room = String(room || rooms[0].name);
    var all = readRosterRaw_(true);
    if (baseRev && roomRevs_(all, rooms)[room] !== baseRev) {
      // มีคนอื่นแก้ไขรายชื่อห้องนี้ไปก่อน (เพิ่ม/ลบ/แก้ไข) — ไม่บันทึกทับ ส่งรายชื่อล่าสุดกลับให้หน้าเว็บตัดสินใจ
      return { success: false, conflict: true, students: getRoster_(true), revs: roomRevs_(all, rooms),
               message: 'รายชื่อห้องนี้ถูกแก้ไขโดยผู้ใช้อื่นระหว่างที่คุณกำลังแก้ไข จึงยังไม่ได้บันทึก (กันข้อมูลทับกัน)' };
    }
    var maxNum = maxId_(all);
    (rows || []).forEach(function (r) {
      var n = parseInt(r.id, 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    });
    var seen = {};
    all.forEach(function (s) { if (s.room !== room) seen[s.id] = 1; });
    var mine = [];
    (rows || []).forEach(function (r) {
      var name = String(r.name || '').trim();
      if (!name) return;
      var id = String(r.id || '').trim();
      if (!id || seen[id]) {
        do { maxNum++; id = String(maxNum); } while (seen[id]);
      }
      seen[id] = 1;
      mine.push({
        id: id, name: name, room: room,
        nickname: String(r.nickname || '').trim(), gender: String(r.gender || '').trim(),
        birth: String(r.birth || '').trim(), note: String(r.note || '').trim()
      });
    });
    var list = all.filter(function (s) { return s.room !== room; }).concat(mine);
    writeRoster_(o.roster, list);
    return { success: true, message: 'บันทึกรายชื่อนักเรียนเรียบร้อยแล้ว', students: getRoster_(true), revs: roomRevs_(readRosterRaw_(true), rooms) };
  }, ['roster']);
}

// ---------- ห้องเรียน ----------
var LEVELS = ['อ.1', 'อ.2', 'อ.3', 'ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6'];

// อ่านชื่อห้องได้หลายแบบ: อ.2 / อ2 / อนุบาล 2 / ชั้นอนุบาลปีที่ 2 / ป.3/1 / ประถมศึกษาปีที่ 3 ห้อง 1
// คืน {level, suffix, key} หรือ null ถ้าอ่านไม่ออก (ตัวเลขล้วน เช่น "3/1" ถือว่าอ่านไม่ออก)
function parseRoom_(t) {
  t = String(t == null ? '' : t).replace(/\s+/g, ' ').trim().replace(/^ชั้น\s*/, '');
  if (!t) return null;
  var m = t.match(/^(อนุบาลศึกษา|อนุบาล|ประถมศึกษา|ประถม|อ|ป)\s*\.?\s*(?:ปีที่\s*)?([1-6])(?![0-9])\s*(.*)$/);
  if (!m) return null;
  var pre = m[1].charAt(0) === 'อ' ? 'อ' : 'ป';
  var n = Number(m[2]);
  if (pre === 'อ' && n > 3) return null;
  var suf = String(m[3] || '').replace(/^\s*(?:\/|-|ห้องที่|ห้อง|\.)\s*/, '').trim();
  if (suf.length > 12) return null;
  var level = pre + '.' + n;
  return { level: level, suffix: suf, key: level + (suf ? '/' + suf : '') };
}
function levelFromName_(t) { var p = parseRoom_(t); return p ? p.level : ''; }
function levelFull_(level) { return (level.charAt(0) === 'อ' ? 'อนุบาลปีที่ ' : 'ประถมศึกษาปีที่ ') + level.slice(2); }
function roomFull_(key) {
  var p = parseRoom_(key);
  return p ? 'ชั้น' + levelFull_(p.level) + (p.suffix ? '/' + p.suffix : '') : String(key);
}
function autoPos_(key) { return 'ครูประจำชั้น' + roomFull_(key).replace(/^ชั้น/, ''); }

function ensureRooms_(ss) {
  var sh = ss.getSheetByName(SH.rooms);
  if (sh) return sh;
  sh = ss.insertSheet(SH.rooms);
  sh.getRange(1, 1, 1, 4).setValues([['ชื่อห้อง', 'ระดับชั้น', 'ครูประจำชั้น', 'ตำแหน่งครู']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // ย้ายข้อมูลจากเวอร์ชันเก่า (1 ห้อง): สร้างห้องจากค่าตั้งค่าเดิม
  var st = getSettings_();
  var g = String(st.grade || '').trim();
  var p = parseRoom_(g);
  var key = p ? p.key : 'อ.2', lv = p ? p.level : 'อ.2';
  sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
  sh.getRange(2, 1, 1, 4).setValues([[key, lv, st.teacherName || '', st.teacherPos || autoPos_(key)]]);
  return sh;
}

// แปลงชื่อห้องเดิม (เช่น "อนุบาลศึกษาปีที่ 2") เป็นรูปแบบมาตรฐาน (อ.2) พร้อมแก้ชื่อห้องในรายชื่อนักเรียน
function normalizeRooms_(o) {
  var sh = o.rooms, last = sh.getLastRow();
  if (last < 2) return;
  var v = sh.getRange(2, 1, last - 1, 4).getValues(), ren = {}, changed = false, taken = {};
  v.forEach(function (r) { taken[String(r[0]).trim()] = 1; });
  v.forEach(function (r) {
    var name = String(r[0]).trim();
    if (!name) return;
    var p = parseRoom_(name);
    if (!p) return;
    var lv = String(r[1]).trim();
    var level = LEVELS.indexOf(lv) >= 0 ? lv : p.level;
    var key = level + (p.suffix ? '/' + p.suffix : '');
    if (key !== name && !taken[key]) { ren[name] = key; r[0] = key; taken[key] = 1; changed = true; }
    if (String(r[1]).trim() !== level) { r[1] = level; changed = true; }
  });
  if (!changed) return;
  sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
  sh.getRange(2, 1, v.length, 4).setValues(v);
  var rs = o.roster, h = rosterHeaderRow_(rs), rl = rs.getLastRow();
  if (rl > h && Object.keys(ren).length) {
    var rg = rs.getRange(h + 1, 8, rl - h, 1), rv = rg.getValues();
    rg.setNumberFormat('@');
    rg.setValues(rv.map(function (x) { var t = String(x[0]).trim(); return [ren.hasOwnProperty(t) ? ren[t] : x[0]]; }));
  }
}

function loadRooms_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureRooms_(ss);
  var last = sh.getLastRow();
  var out = [];
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) {
      var name = String(r[0]).trim();
      if (!name) return;
      var lvl = String(r[1]).trim() || levelFromName_(name) || 'อ.2';
      out.push({ name: name, level: lvl, teacher: String(r[2]).trim(),
                 pos: String(r[3]).trim() || autoPos_(name),
                 suffix: (parseRoom_(name) || { suffix: '' }).suffix, full: roomFull_(name) });
    });
  }
  if (!out.length) {
    sh.getRange(2, 1, 1, 4).setValues([['อ.2', 'อ.2', '', autoPos_('อ.2')]]);
    out.push({ name: 'อ.2', level: 'อ.2', teacher: '', pos: autoPos_('อ.2'), suffix: '', full: roomFull_('อ.2') });
  }
  return out;
}

// rooms: [{name, old (ชื่อเดิม ถ้าเปลี่ยนชื่อ), level, teacher, pos}]
function saveRooms_(rooms) {
  return withLock_(function () {
    var o = ensureSetup_();
    var clean = [], seen = {};
    (rooms || []).forEach(function (r) {
      var lv = String(r.level || '').trim();
      var sfx = String(r.suffix == null ? '' : r.suffix).trim().replace(/^\//, '');
      var name;
      if (LEVELS.indexOf(lv) >= 0) name = lv + (sfx ? '/' + sfx : '');
      else { var pr = parseRoom_(r.name); if (!pr) return; lv = pr.level; name = pr.key; }
      if (seen[name]) return;
      seen[name] = 1;
      clean.push({ name: name, old: String(r.old || name), level: lv,
                   teacher: String(r.teacher || '').trim(), pos: String(r.pos || '').trim() || autoPos_(name) });
    });
    if (!clean.length) return { success: false, message: 'ต้องมีอย่างน้อย 1 ห้อง (ตรวจชื่อห้องซ้ำ)' };
    var all = readRosterRaw_(true);
    var keepNames = {};
    clean.forEach(function (r) { keepNames[r.old] = r.name; });
    var rmap = {}; // เปลี่ยนชื่อ/ลบห้อง -> สิทธิ์ห้องของผู้ใช้ตามไปด้วย
    getRooms_(true).forEach(function (r) { rmap[r.name] = Object.prototype.hasOwnProperty.call(keepNames, r.name) ? keepNames[r.name] : null; });
    var orphan = all.filter(function (s) { return !keepNames.hasOwnProperty(s.room); });
    if (orphan.length) {
      return { success: false, message: 'ลบห้องไม่ได้ เพราะยังมีนักเรียนอยู่ ' + orphan.length + ' คน (ย้ายหรือลบนักเรียนออกก่อน)' };
    }
    all.forEach(function (s) { s.room = keepNames[s.room]; });
    writeRoster_(o.roster, all);
    var sh = o.rooms;
    var last = sh.getLastRow();
    sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
    if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
    sh.getRange(2, 1, clean.length, 4).setValues(clean.map(function (r) { return [r.name, r.level, r.teacher, r.pos]; }));
    remapUserRooms_(rmap);
    return { success: true, message: 'บันทึกห้องเรียนเรียบร้อยแล้ว', rooms: getRooms_(true), students: getRoster_(true),
             revs: roomRevs_(readRosterRaw_(true), getRooms_(true)) };
  }, ['meta', 'roster', 'users']);
}

// ---------- เลื่อนชั้นสู่ปีการศึกษาใหม่ ----------
function nextLevel_(lv) {
  var i = LEVELS.indexOf(lv);
  if (i < 0) return lv;
  return i + 1 < LEVELS.length ? LEVELS[i + 1] : null; // null = จบ ป.6
}

// คืนแผนการเลื่อนชั้น (ใช้ทั้งแสดงตัวอย่างและทำจริง)
function promotePlan_(rooms) {
  var plan = rooms.map(function (r) {
    var nl = nextLevel_(r.level);
    var p = parseRoom_(r.name);
    var sfx = p ? p.suffix : '';
    var nn = nl ? nl + (sfx ? '/' + sfx : '') : r.name;
    return { old: r.name, level: r.level, newLevel: nl, newName: nn, grad: !nl };
  });
  var seen = {};
  plan.forEach(function (p) { if (!p.grad) { if (seen[p.newName]) p.newName = p.old; seen[p.newName] = 1; } });
  return plan;
}

// fromYear = ปีการศึกษาที่หน้าเว็บเห็นอยู่ (กันกดซ้ำ) / removeGraduates = ลบรายชื่อนักเรียน ป.6
function promoteYear_(fromYear, removeGraduates) {
  return withLock_(function () {
    var o = ensureSetup_();
    var st = getSettings_(true);
    if (Number(fromYear) !== Number(st.academicYear)) {
      return { success: false, message: 'ปีการศึกษาในระบบเปลี่ยนไปแล้ว (ตอนนี้ ' + st.academicYear + ') โหลดหน้าใหม่แล้วตรวจอีกครั้ง' };
    }
    var rooms = getRooms_(true);
    var plan = promotePlan_(rooms);
    var all = readRosterRaw_(true);
    var map = {};
    plan.forEach(function (p) { map[p.old] = p; });
    var removed = 0;
    var kept = [];
    all.forEach(function (x) {
      var p = map[x.room];
      if (p && p.grad && removeGraduates) { removed++; return; }
      if (p) x.room = p.newName;
      kept.push(x);
    });
    var newRooms = [];
    plan.forEach(function (p, i) {
      if (p.grad && removeGraduates) return;
      var r = rooms[i];
      var pos = (r.pos === autoPos_(r.name)) ? autoPos_(p.newName) : r.pos; // ตำแหน่งที่เติมอัตโนมัติ เปลี่ยนตามชั้นใหม่
      newRooms.push([p.newName, p.grad ? r.level : p.newLevel, r.teacher, p.grad ? r.pos : pos]);
    });
    if (!newRooms.length) newRooms.push(['อ.1', 'อ.1', '', 'ครูประจำชั้น']); // ไม่เหลือห้องเลย: เปิดห้องรับนักเรียนใหม่
    var rs = o.rooms, last = rs.getLastRow();
    rs.getRange(1, 1, rs.getMaxRows(), 1).setNumberFormat('@');
    if (last > 1) rs.getRange(2, 1, last - 1, 4).clearContent();
    rs.getRange(2, 1, newRooms.length, 4).setValues(newRooms);
    var rmap = {};
    plan.forEach(function (p) { rmap[p.old] = (p.grad && removeGraduates) ? null : p.newName; });
    remapUserRooms_(rmap);
    writeRoster_(o.roster, kept);
    // ปีการศึกษา +1 และกลับไปเดือนเริ่มต้น
    var ss = o.settings, sl = ss.getLastRow();
    var vals = sl >= 3 ? ss.getRange(3, 1, sl - 2, 3).getValues() : [];
    vals.forEach(function (r, i) {
      if (String(r[2]) === 'academicYear') ss.getRange(i + 3, 2).setValue(Number(st.academicYear) + 1);
      if (String(r[2]) === 'currentMonth') ss.getRange(i + 3, 2).setValue(Number(st.startMonth));
    });
    return { success: true, message: 'เลื่อนชั้นเรียบร้อยแล้ว ปีการศึกษา ' + (Number(st.academicYear) + 1),
             removed: removed, settings: getSettings_(true), rooms: getRooms_(true), students: getRoster_(true),
             revs: roomRevs_(readRosterRaw_(true), getRooms_(true)) };
  }, ['meta', 'roster', 'users', 'data']);
}

// ---------- นำเข้ารายชื่อ ----------
// items: [{room,id,name,nickname,gender,birth,note,w,h}] (หน้าเว็บแปลงและตรวจรูปแบบมาแล้ว)
// กติกา: จับคู่ด้วยรหัสประจำตัวก่อน ไม่มีจึงจับด้วยชื่อ-สกุล -> อัปเดตเฉพาะช่องที่มีค่า / ไม่พบ -> เพิ่มใหม่ / ไม่ลบใคร
function importStudents_(items, defaultRoom, year, month, ctx) {
  return withLock_(function () {
    var o = ensureSetup_();
    var rooms = getRooms_(true);
    var roomSet = {};
    rooms.forEach(function (r) { roomSet[r.name] = 1; });
    var all = readRosterRaw_(true);
    var byId = {}, byName = {};
    // ครู: จับคู่ด้วยชื่อเฉพาะนักเรียนในห้องที่ตนมีสิทธิ์ (ไม่ไปแก้/ย้ายนักเรียนห้องอื่น)
    all.forEach(function (s) {
      byId[s.id] = s;
      if (!ctx || ctx.allRooms || roomAllowed_(ctx, s.room)) byName[normName_(s.name)] = s;
    });
    var maxNum = maxId_(all);
    var added = 0, updated = 0, denied = 0, newRooms = [], hw = [];
    (items || []).forEach(function (it) {
      var name = String(it.name || '').trim();
      if (!name) return;
      var room = String(it.room || defaultRoom || rooms[0].name).trim();
      var pr = parseRoom_(room);
      if (pr) room = pr.key;
      if (ctx && !ctx.allRooms && !roomAllowed_(ctx, room)) { denied++; return; }
      if (!roomSet[room] && ctx && ctx.role !== ROLE_ADMIN) { denied++; return; } // เฉพาะ Admin สร้างห้องใหม่ได้
      if (!roomSet[room]) {
        var lv = pr ? pr.level : 'อ.2';
        o.rooms.getRange(1, 1, o.rooms.getMaxRows(), 1).setNumberFormat('@');
        o.rooms.getRange(o.rooms.getLastRow() + 1, 1, 1, 4).setValues([[room, lv, '', autoPos_(room)]]);
        roomSet[room] = 1; newRooms.push(room);
      }
      var id = String(it.id || '').trim();
      var st = (id && byId[id]) || byName[normName_(name)];
      if (st && ctx && !ctx.allRooms && !roomAllowed_(ctx, st.room)) { denied++; return; }
      if (st) {
        if (it.room) st.room = room;
        st.name = name;
        ['nickname', 'gender', 'birth', 'note'].forEach(function (k) {
          var v = String(it[k] || '').trim();
          if (v) st[k] = v;
        });
        updated++;
      } else {
        if (!id || byId[id]) { do { maxNum++; id = String(maxNum); } while (byId[id]); }
        else { var nn = parseInt(id, 10); if (!isNaN(nn) && nn > maxNum) maxNum = nn; }
        st = { id: id, name: name, room: room, nickname: String(it.nickname || '').trim(),
               gender: String(it.gender || '').trim(), birth: String(it.birth || '').trim(),
               note: String(it.note || '').trim() };
        all.push(st); byId[id] = st; byName[normName_(name)] = st;
        added++;
      }
      var w = parseFloat(it.w), h = parseFloat(it.h);
      if (!isNaN(w) || !isNaN(h)) hw.push({ id: st.id, w: w, h: h });
    });
    writeRoster_(o.roster, all);
    var hwCount = 0;
    year = Number(year); month = Number(month);
    if (hw.length && year && month) {
      // เขียนเฉพาะแถวน้ำหนัก/ส่วนสูงของเดือนนี้ที่เปลี่ยนหรือเพิ่ม (ช่องที่ไฟล์ไม่มีค่า คงค่าเดิมในชีต)
      var sh = o.health, last = sh.getLastRow();
      var hdata = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
      var patches = {};
      hw.forEach(function (x) {
        var p = patches[x.id] || (patches[x.id] = {});
        if (!isNaN(x.w)) p.w = x.w;
        if (!isNaN(x.h)) p.h = x.h;
        hwCount++;
      });
      var plan = planHealth_(hdata, year, month, patches, null);
      applyRowOps_(sh, 5, 4, hdata, plan.updates, plan.appends, plan.removes);
    }
    return { success: true, message: 'นำเข้าเรียบร้อย', added: added, updated: updated, denied: denied,
             newRooms: newRooms, health: hwCount, rooms: getRooms_(true), students: getRoster_(true),
             revs: roomRevs_(readRosterRaw_(true), getRooms_(true)) };
  }, ['meta', 'roster', 'data']);
}

function normName_(n) { return String(n || '').replace(/\s+/g, ' ').trim(); }

// รหัสสูงสุดที่เคยถูกใช้ — นับทั้งรายชื่อและข้อมูลบันทึกทุกชีต เพื่อไม่ให้นักเรียนใหม่ได้รหัสของคนเก่า
function maxId_(list) {
  var m = 1000;
  list.forEach(function (s) { var n = parseInt(s.id, 10); if (!isNaN(n) && n > m) m = n; });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [[SH.routine, 4], [SH.health, 3], [SH.check, 3], [SH.indiv, 1]].forEach(function (d) {
    var sh = ss.getSheetByName(d[0]);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, d[1], sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var n = parseInt(r[0], 10);
      if (!isNaN(n) && n > m) m = n;
    });
  });
  return m;
}

// อ่านรายชื่อทั้งหมด (ทุกห้อง) พร้อมคอลัมน์ห้อง; ห้องที่ว่าง = ห้องแรก
function loadRosterRaw_(fresh) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureRoster_(ss);
  var h = rosterHeaderRow_(sh);
  var last = sh.getLastRow();
  if (last <= h) return [];
  var tz = ss.getSpreadsheetTimeZone();
  var first = getRoomsQuick_(fresh);
  var out = [];
  sh.getRange(h + 1, 1, last - h, 8).getValues().forEach(function (r) {
    var name = String(r[2]).trim();
    if (!name) return;
    out.push({
      id: String(r[1]).trim() || ('S' + (out.length + 1)),
      name: name, nickname: String(r[3]).trim(), gender: String(r[4]).trim(),
      birth: fmtDate_(r[5], tz), note: String(r[6]).trim(),
      room: String(r[7]).trim() || first
    });
  });
  return out;
}

function getRoomsQuick_(fresh) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.rooms);
  if (!sh || sh.getLastRow() < 2) return getRooms_(fresh)[0].name;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim()) return String(v[i][0]).trim();
  return getRooms_(fresh)[0].name;
}

// เขียนรายชื่อทั้งหมดกลับ จัดเรียงตามลำดับห้อง และเลขที่รันใหม่ในแต่ละห้อง
function writeRoster_(sh, list) {
  var h = rosterHeaderRow_(sh);
  var order = getRooms_(true).map(function (r) { return r.name; });
  var groups = {};
  list.forEach(function (s) { (groups[s.room] = groups[s.room] || []).push(s); });
  var names = order.filter(function (n) { return groups[n]; });
  Object.keys(groups).forEach(function (n) { if (names.indexOf(n) < 0) names.push(n); });
  var rows = [];
  names.forEach(function (n) {
    groups[n].forEach(function (s, i) {
      rows.push([i + 1, s.id, s.name, s.nickname || '', s.gender || '', s.birth || '', s.note || '', s.room]);
    });
  });
  var last = sh.getLastRow();
  if (last > h) sh.getRange(h + 1, 1, last - h, 8).clearContent();
  if (rows.length) {
    ensureRows_(sh, h + rows.length);
    [2, 3, 4, 5, 6, 7, 8].forEach(function (c) { sh.getRange(h + 1, c, rows.length, 1).setNumberFormat('@'); });
    sh.getRange(h + 1, 1, rows.length, 8).setValues(rows);
  }
}

function saveHolidays_(rows) {
  return withLock_(function () {
    var sh = ensureSetup_().holidays;
    var h = holidayHeaderRow_(sh);
    var clean = (rows || []).filter(function (r) { return /^\d{4}-\d{2}-\d{2}$/.test(String(r.date || '')); })
      .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; })
      .map(function (r) {
        var p = r.date.split('-');
        var dow = THAI_DOW[new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getDay()];
        return ['', r.date, 'วัน' + dow, String(r.name || ''), r.open ? true : false];
      });
    var last = sh.getLastRow();
    if (last > h) sh.getRange(h + 1, 1, last - h, 5).clearContent();
    if (clean.length) {
      sh.getRange(h + 1, 2, clean.length, 1).setNumberFormat('@');
      sh.getRange(h + 1, 1, clean.length, 5).setValues(clean);
    }
    return { success: true, message: 'บันทึกวันหยุดเรียบร้อยแล้ว', holidays: getHolidays_(true) };
  }, ['meta']);
}

/* ------------------------------------------------------------------ */
/*  โครงสร้างชีต                                                        */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/*  แคชฝั่งเซิร์ฟเวอร์ (CacheService) — เพิ่มความเร็ว โดยข้อมูลต้นฉบับอยู่ในชีตเสมอ        */
/*                                                                      */
/*  1) แคชเป็นแค่ "สำเนา": ถ้าแคชหาย/หมดอายุ/ใช้ไม่ได้ ระบบอ่านจากชีตใหม่เองทุกครั้ง      */
/*  2) แบ่งเป็น 4 กลุ่ม: users / meta (ตั้งค่า ห้อง วันหยุด เกณฑ์โภชนาการ) / roster / data */
/*     (นม แปรงฟัน น้ำหนัก-ส่วนสูง ตรวจสุขภาพ) แต่ละกลุ่มมี "เลขรุ่น" (gen)               */
/*     ทุกครั้งที่ฟังก์ชันบันทึกข้อมูลเสร็จจะเปลี่ยนเลขรุ่น (bump_) แคชรุ่นเก่าจึงถูกทิ้งทันที  */
/*  3) โค้ดที่เขียนข้อมูล (อยู่ใน withLock_) ต้องอ่านแบบ fresh=true เสมอ ห้ามอ่านจากแคช   */
/*     และข้อมูลที่ส่งกลับให้หน้าเว็บหลังเขียนก็ต้อง fresh=true เช่นกัน                   */
/*  4) ถ้าแก้ชีตด้วยมือ แคชจะหมดอายุเองตาม TTL หรือสั่งเมนู "ล้างแคชระบบ" ได้ทันที        */
/* ------------------------------------------------------------------ */
var CACHE_GROUPS = ['users', 'meta', 'roster', 'data'];
var CACHE_TTL = { users: 300, settings: 300, rooms: 300, roster: 300, holidays: 600, criteria: 1800, years: 120, period: 120 };
var CACHE_CHUNK = 30000;   // ตัวอักษรต่อชิ้น (ภาษาไทย 3 ไบต์/ตัว => ไม่เกิน ~90KB ต่ำกว่าขีดจำกัด 100KB ต่อค่า)
var _gens = null;          // เลขรุ่นของแต่ละกลุ่ม (จำไว้ตลอดการประมวลผลหนึ่งครั้ง)
var _mem = {};             // ค่าที่อ่านแล้วในการประมวลผลหนึ่งครั้ง (หายเองเมื่อจบคำขอ)

function newGen_() { return randHex_().substring(0, 12); }

function gens_() {
  if (_gens) return _gens;
  var out = {};
  try {
    var c = CacheService.getScriptCache();
    var got = c.getAll(CACHE_GROUPS.map(function (g) { return 'G:' + g; })), put = {}, need = false;
    CACHE_GROUPS.forEach(function (g) {
      var v = got['G:' + g];
      if (!v) { v = newGen_(); put['G:' + g] = v; need = true; }
      out[g] = v;
    });
    if (need) c.putAll(put, 21600);
  } catch (e) {
    // แคชใช้ไม่ได้: ใช้เลขสุ่ม = แคชไม่มีวัน hit ระบบจะอ่านจากชีตทุกครั้ง
    CACHE_GROUPS.forEach(function (g) { out[g] = 'x' + newGen_(); });
  }
  _gens = out;
  return out;
}

// เปลี่ยนเลขรุ่นของกลุ่มที่ระบุ (ไม่ระบุ = ทุกกลุ่ม) => แคชเดิมของกลุ่มนั้นใช้ไม่ได้ทันที
function bump_(groups) {
  _mem = {}; _gens = null;
  try {
    var put = {};
    (groups || CACHE_GROUPS).forEach(function (g) { put['G:' + g] = newGen_(); });
    CacheService.getScriptCache().putAll(put, 21600);
  } catch (e) { /* แคชใช้ไม่ได้ = ไม่มีแคชให้ล้าง */ }
}

function chunkPut_(c, key, obj, ttl) {
  var s = JSON.stringify(obj), n = Math.max(1, Math.ceil(s.length / CACHE_CHUNK)), put = {};
  for (var i = 0; i < n; i++) put[key + ':' + i] = (i === 0 ? n + '|' : '') + s.substr(i * CACHE_CHUNK, CACHE_CHUNK);
  c.putAll(put, ttl);
}
function chunkGet_(c, key) {
  var first = c.get(key + ':0');
  if (first === null || first === undefined) return null;
  var bar = first.indexOf('|'), n = Number(first.substring(0, bar)), s = first.substring(bar + 1);
  if (bar < 0 || !(n >= 1)) return null;
  if (n > 1) {
    var keys = [];
    for (var i = 1; i < n; i++) keys.push(key + ':' + i);
    var got = c.getAll(keys);
    for (var j = 1; j < n; j++) {
      var p = got[key + ':' + j];
      if (p === null || p === undefined) return null;   // ชิ้นหาย = ถือว่าแคชเสีย อ่านจากชีตแทน
      s += p;
    }
  }
  try { return JSON.parse(s); } catch (e) { return null; }
}

// อ่านค่าผ่านแคช: groups = กลุ่มที่ค่านี้ขึ้นอยู่กับ, name = ชื่อรายการ, loader = ฟังก์ชันอ่านจากชีตจริง
// fresh=true => อ่านจากชีตเสมอ (ใช้ในโค้ดที่เขียนข้อมูล) ห้ามแก้ไขค่าที่ได้จากฟังก์ชันนี้โดยไม่สำเนาก่อน
function cached_(groups, name, ttl, loader, fresh) {
  if (fresh) { delete _mem[name]; return loader(); }
  if (Object.prototype.hasOwnProperty.call(_mem, name)) return _mem[name];
  var key = null, c = null;
  try {
    var g = gens_();
    key = 'D:' + groups.map(function (x) { return g[x]; }).join('.') + ':' + name;
    c = CacheService.getScriptCache();
    var hit = chunkGet_(c, key);
    if (hit && Object.prototype.hasOwnProperty.call(hit, 'v')) { _mem[name] = hit.v; return hit.v; }
  } catch (e) { key = null; }
  var v = loader();
  _mem[name] = v;
  if (key) { try { chunkPut_(c, key, { v: v }, ttl); } catch (e2) { /* เก็บแคชไม่ได้ ไม่กระทบการทำงาน */ } }
  return v;
}

// ล็อกเขียนข้อมูล (ทีละคน) — groups = กลุ่มแคชที่ฟังก์ชันนี้แก้ไข จะถูก bump หลัง flush ก่อนปล่อยล็อกเสมอ
// (แม้เกิด error กลางทาง ก็ล้างแคชเพื่อกันข้อมูลค้าง) ห้ามเรียก withLock_ ซ้อนกัน
function withLock_(fn, groups) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); }
  finally {
    try { if (groups && groups.length) { SpreadsheetApp.flush(); bump_(groups); } } catch (e) { /* ปล่อยล็อกต่อไป */ }
    lock.releaseLock();
  }
}

function ensureSheet_(ss, name, headers, textCols) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    (textCols || []).forEach(function (c) {
      sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  }
  return sh;
}

// ตรวจโครงสร้างชีต "เต็มรูปแบบ" (สร้างชีตที่ขาด/ย้ายข้อมูลเวอร์ชันเก่า) ใช้เมื่อ:
//  (1) แคชธงยืนยันโครงสร้างหมดอายุ (ทุก ~10 นาที หรือหลังขึ้นเวอร์ชันใหม่ SETUP_VERSION)
//  (2) สั่งบังคับ (เมนู "ตรวจ/สร้างโครงสร้างชีต")  (3) พบว่าชีตที่ต้องใช้หายไปตอนเปิดใช้
// ถ้าแก้โครงสร้างชีต/คอลัมน์ในอนาคต ให้เปลี่ยนเลข SETUP_VERSION — ระบบจะตรวจเต็มรูปแบบอีกครั้งทันที
var SETUP_VERSION = '2026-10-31.1';
var SETUP_TTL = 600;
var _setup = null;   // ตัวจับชีต จำไว้ตลอดการประมวลผลหนึ่งครั้ง

function ensureSetup_(force) {
  if (!force && _setup) return _setup;
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ok = false;
  if (!force) {
    try { ok = CacheService.getScriptCache().get('SETUP_OK:' + SETUP_VERSION) === '1'; } catch (e) { ok = false; }
  }
  if (!ok) {
    ensureSetupFull_(ss);
    try { CacheService.getScriptCache().put('SETUP_OK:' + SETUP_VERSION, '1', SETUP_TTL); } catch (e2) { /* ไม่มีแคช = ตรวจเต็มรูปแบบทุกครั้ง (ช้าแต่ปลอดภัย) */ }
    bump_(['meta', 'roster', 'users']);   // ตรวจเต็มรูปแบบอาจแก้ชื่อห้อง/ตั้งค่า จึงล้างแคชกลุ่มที่เกี่ยวข้อง
  }
  _setup = lazySheets_(ss);
  return _setup;
}

// ตัวจับชีตแบบ lazy: เปิดเฉพาะชีตที่ใช้จริง และถ้าชีตนั้นหายไป (เช่นถูกลบด้วยมือ)
// จะสร้างโครงสร้างใหม่ให้ทันที ไม่ปล่อยให้ระบบมองข้ามชีตที่จำเป็น
function lazySheets_(ss) {
  var o = {};
  Object.keys(SH).forEach(function (k) {
    var memo = null;
    Object.defineProperty(o, k, {
      enumerable: true,
      get: function () {
        if (memo) return memo;
        var sh = ss.getSheetByName(SH[k]);
        if (!sh) { sh = ensureSetupFull_(ss)[k]; bump_(); }
        memo = sh;
        return sh;
      }
    });
  });
  return o;
}

// สร้างชีตที่ขาด (ทำซ้ำได้ ไม่ทับข้อมูลเดิม) และคืนอ้างอิงชีตทั้งหมด
function ensureSetupFull_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var o = {};
  o.settings = ensureSettings_(ss);
  o.roster = ensureRoster_(ss);
  o.rooms = ensureRooms_(ss);
  normalizeRooms_(o);
  o.holidays = ensureHolidays_(ss);
  o.criteria = ensureCriteria_(ss);
  o.routine = ensureSheet_(ss, SH.routine,
    ['ปีการศึกษา', 'เดือน', 'ประเภท', 'รหัสนักเรียน', 'วันที่ที่ทำ', 'รวม (วัน)'], [3, 4, 5]);
  o.health = ensureSheet_(ss, SH.health,
    ['ปีการศึกษา', 'เดือน', 'รหัสนักเรียน', 'น้ำหนัก (กก.)', 'ส่วนสูง (ซม.)'], [3]);
  o.check = ensureSheet_(ss, SH.check,
    ['ปีการศึกษา', 'เดือน', 'รหัสนักเรียน', 'ผลตรวจ (20 ช่อง: 4 สัปดาห์ x เล็บ/ฟัน/ผม/ร่างกาย + สรุป; g=ดี x=ปรับปรุง -=ว่าง)'], [3, 4]);
  o.indiv = ensureSheet_(ss, SH.indiv, ['รหัสนักเรียน', 'ข้อมูล (JSON)', 'แก้ไขล่าสุด'], [1, 2]);
  o.users = ensureUsers_(ss);
  return o;
}

/* ---- ตัวอ่านข้อมูลผ่านแคช (ชื่อเดิม พารามิเตอร์ fresh=true = อ่านจากชีตจริง ใช้ในโค้ดที่เขียนข้อมูล) ---- */
function getSettings_(fresh) { return cached_(['meta'], 'settings', CACHE_TTL.settings, loadSettings_, fresh); }
function getRooms_(fresh) { return cached_(['meta'], 'rooms', CACHE_TTL.rooms, loadRooms_, fresh); }
function readRosterRaw_(fresh) {
  // รายชื่อ "ห้องที่ว่าง = ห้องแรก" ขึ้นกับชีตห้องเรียน จึงขึ้นกับกลุ่ม meta ด้วย
  return cached_(['roster', 'meta'], 'roster', CACHE_TTL.roster, function () { return loadRosterRaw_(!!fresh); }, fresh);
}
function getHolidays_(fresh) { return cached_(['meta'], 'holidays', CACHE_TTL.holidays, loadHolidays_, fresh); }
function getCriteria_(fresh) { return cached_(['meta'], 'criteria', CACHE_TTL.criteria, loadCriteria_, fresh); }
function getYears_(fresh) { return cached_(['data'], 'years', CACHE_TTL.years, loadYears_, fresh); }
function readUsers_(fresh) { return cached_(['users'], 'users', CACHE_TTL.users, loadUsers_, fresh); }

// ตารางเกณฑ์น้ำหนักตามเกณฑ์ส่วนสูง (กรมอนามัย): น้ำหนัก (กก.) ที่เส้น SD แต่ละเส้น ณ ส่วนสูงนั้น
var CRIT_HEADERS = ['เพศ (ชาย/หญิง)', 'ส่วนสูง (ซม.)', '-3SD', '-2SD', '-1.5SD', 'มัธยฐาน', '+1.5SD', '+2SD', '+3SD'];

// ตารางเกณฑ์น้ำหนักตามเกณฑ์ส่วนสูง เด็กอายุ 2-5 ปี (ส่วนสูง 65-140 ซม. ทีละ 1 ซม.) ดึงจากกราฟของกรมอนามัย พ.ศ. 2564
var CRIT_SEED = {
  'ชาย': {
    m2: [6.34,6.54,6.73,6.93,7.12,7.32,7.5,7.69,7.86,8.03,8.21,8.37,8.54,8.68,8.84,9,9.17,9.36,9.54,9.75,9.96,10.17,10.39,10.61,10.82,11.02,11.22,11.43,11.63,11.83,12.02,12.23,12.43,12.65,12.86,13.09,13.32,13.55,13.8,14.04,14.29,14.54,14.8,15.06,15.33,15.61,15.9,16.18,16.49,16.79,17.09,17.39,17.71,18.01,18.32,18.62,18.96,19.3,19.63,19.96,20.29,20.62,20.97,21.3,21.63,21.96,22.33,22.7,23.1,23.48,23.9,24.35,24.78,25.22,25.68,26.13],
    m15: [6.59,6.8,7,7.2,7.4,7.6,7.8,7.99,8.18,8.36,8.54,8.71,8.88,9.04,9.2,9.37,9.54,9.73,9.92,10.13,10.35,10.58,10.81,11.02,11.24,11.46,11.67,11.87,12.07,12.29,12.49,12.71,12.92,13.15,13.38,13.61,13.86,14.1,14.36,14.62,14.88,15.15,15.42,15.69,15.98,16.28,16.59,16.89,17.2,17.52,17.83,18.17,18.5,18.82,19.14,19.47,19.84,20.19,20.55,20.91,21.27,21.63,21.99,22.35,22.71,23.07,23.47,23.87,24.29,24.73,25.19,25.67,26.14,26.63,27.12,27.62],
    p15: [8.43,8.7,8.95,9.21,9.47,9.73,9.99,10.25,10.49,10.72,10.95,11.18,11.4,11.6,11.8,12.01,12.23,12.46,12.7,12.96,13.22,13.49,13.77,14.04,14.31,14.57,14.85,15.11,15.36,15.64,15.91,16.2,16.48,16.79,17.1,17.43,17.77,18.11,18.45,18.83,19.2,19.57,19.96,20.36,20.78,21.2,21.63,22.07,22.54,22.99,23.45,23.93,24.4,24.88,25.37,25.84,26.48,27.09,27.72,28.34,28.97,29.59,30.23,30.84,31.47,32.1,32.81,33.53,34.28,35.05,35.85,36.68,37.52,38.35,39.21,40.05],
    p2: [8.8,9.08,9.36,9.63,9.9,10.17,10.43,10.69,10.95,11.19,11.44,11.68,11.9,12.12,12.33,12.54,12.78,13.02,13.26,13.54,13.81,14.08,14.37,14.64,14.93,15.2,15.48,15.75,16.04,16.31,16.6,16.9,17.2,17.52,17.85,18.19,18.5,18.91,19.3,19.68,20.08,20.48,20.9,21.31,21.76,22.21,22.68,23.15,23.63,24.11,24.61,25.12,25.62,26.14,26.6,27.17,27.87,28.57,29.26,29.95,30.65,31.34,32.05,32.74,33.44,34.14,34.92,35.72,36.55,37.39,38.28,39.19,40.11,41.04,41.97,42.88],
    p3: [9.72,10.02,10.32,10.62,10.92,11.21,11.51,11.8,12.07,12.3,12.62,12.88,13.12,13.36,13.59,13.82,14.07,14.33,14.6,14.87,15.17,15.47,15.78,16.08,16.38,16.68,16.97,17.28,17.58,17.89,18.21,18.54,18.87,19.23,19.6,19.99,20.38,20.79,21.22,21.66,22.1,22.56,23.04,23.52,24.01,24.53,25.06,25.59,26.16,26.71,27.27,27.85,28.44,29.03,29.62,30.21,31.09,31.98,32.85,33.73,34.61,35.58,36.57,37.54,38.52,39.51,40.5,41.48,42.53,43.59,44.69,45.82,46.96,48.09,49.24,50.36]
  },
  'หญิง': {
    m2: [6.07,6.26,6.43,6.62,6.79,6.97,7.14,7.31,7.47,7.64,7.8,7.96,8.1,8.28,8.43,8.61,8.8,8.98,9.18,9.38,9.59,9.8,10.02,10.22,10.43,10.65,10.85,11.05,11.26,11.47,11.67,11.89,12.1,12.32,12.54,12.77,13.01,13.25,13.5,13.77,14.03,14.31,14.59,14.9,15.2,15.52,15.83,16.15,16.48,16.82,17.16,17.5,17.85,18.19,18.53,18.88,19.14,19.4,19.66,19.92,20.16,20.46,20.77,21.07,21.37,21.66,21.99,22.31,22.64,22.95,23.27,23.67,24.08,24.51,24.96,25.45],
    m15: [6.34,6.54,6.72,6.91,7.1,7.27,7.46,7.63,7.8,7.97,8.13,8.29,8.46,8.62,8.79,8.98,9.17,9.37,9.57,9.79,10.01,10.22,10.45,10.67,10.88,11.1,11.31,11.53,11.75,11.96,12.18,12.41,12.62,12.85,13.08,13.33,13.59,13.83,14.1,14.38,14.67,14.95,15.25,15.56,15.89,16.21,16.55,16.9,17.24,17.6,17.96,18.32,18.68,19.04,19.4,19.77,20.06,20.35,20.65,20.94,21.23,21.56,21.89,22.21,22.53,22.87,23.23,23.57,23.93,24.29,24.65,25.08,25.55,26.03,26.53,27.09],
    p15: [8.33,8.59,8.84,9.08,9.31,9.54,9.77,10,10.23,10.45,10.67,10.87,11.07,11.28,11.51,11.74,11.99,12.25,12.51,12.78,13.07,13.35,13.64,13.93,14.22,14.51,14.79,15.08,15.37,15.66,15.95,16.25,16.55,16.87,17.18,17.52,17.86,18.2,18.58,18.95,19.36,19.76,20.17,20.61,21.05,21.52,21.99,22.48,22.97,23.47,23.98,24.48,25,25.51,26.03,26.55,27.12,27.7,28.28,28.85,29.43,30.02,30.61,31.2,31.8,32.39,33.08,33.78,34.46,35.16,35.86,36.66,37.52,38.39,39.32,40.3],
    p2: [8.75,9.01,9.27,9.53,9.77,10.02,10.26,10.51,10.74,10.95,11.18,11.4,11.61,11.85,12.07,12.32,12.56,12.84,13.11,13.4,13.7,14,14.3,14.61,14.91,15.21,15.51,15.81,16.12,16.42,16.74,17.04,17.37,17.69,18.03,18.38,18.74,19.12,19.52,19.92,20.32,20.75,21.2,21.66,22.13,22.63,23.13,23.65,24.16,24.69,25.24,25.79,26.34,26.88,27.44,27.99,28.62,29.28,29.9,30.53,31.17,31.83,32.48,33.14,33.79,34.45,35.23,36.02,36.8,37.57,38.37,39.25,40.18,41.14,42.16,43.22],
    p3: [9.76,10.06,10.35,10.62,10.9,11.17,11.43,11.69,11.95,12.19,12.44,12.68,12.91,13.16,13.41,13.67,13.96,14.2,14.55,14.87,15.2,15.53,15.86,16.21,16.53,16.87,17.2,17.53,17.88,18.22,18.55,18.91,19.27,19.63,20.02,20.41,20.81,21.24,21.67,22.13,22.6,23.08,23.58,24.11,24.65,25.2,25.77,26.35,26.94,27.54,28.16,28.78,29.42,30.03,30.66,31.3,32.07,32.86,33.64,34.42,35.2,36.07,36.96,37.83,38.7,39.59,40.63,41.69,42.72,43.75,44.81,45.89,47,48.15,49.33,50.59]
  }
};

// เด็กอายุ 6-19 ปี (ส่วนสูง 90-185 ซม. ทีละ 1 ซม.) ดึงจากกราฟของกรมอนามัย พ.ศ. 2564
var CRIT_SEED_619 = {
  'ชาย': {
    m2: [11.02,11.23,11.43,11.64,11.83,12.03,12.23,12.43,12.64,12.87,13.1,13.3,13.57,13.8,14.03,14.27,14.53,14.8,15.07,15.34,15.6,15.91,16.17,16.47,16.77,17.07,17.41,17.71,18.01,18.31,18.64,18.97,19.31,19.64,19.98,20.31,20.64,20.98,21.31,21.61,21.94,22.35,22.71,23.08,23.48,23.91,24.35,24.78,25.22,25.68,26.12,26.58,27.05,27.52,27.99,28.45,28.95,29.49,30.02,30.59,31.19,31.79,32.42,33.1,33.72,34.4,35.06,35.69,36.36,37.06,37.77,38.46,39.2,39.9,40.61,41.3,41.98,42.64,43.27,43.94,44.57,45.18,45.8,46.41,46.97,47.55,48.08,48.64,49.18,49.71,50.21,50.74,51.28,51.78,52.29,52.8],
    m15: [11.45,11.67,11.87,12.07,12.27,12.5,12.7,12.93,13.14,13.37,13.6,13.83,14.1,14.34,14.6,14.87,15.14,15.4,15.7,15.97,16.27,16.57,16.87,17.21,17.51,17.84,18.17,18.47,18.81,19.14,19.48,19.84,20.18,20.54,20.91,21.28,21.61,21.98,22.34,22.71,23.04,23.45,23.88,24.29,24.71,25.18,25.65,26.15,26.62,27.11,27.62,28.12,28.62,29.12,29.63,30.16,30.68,31.26,31.82,32.46,33.09,33.76,34.42,35.13,35.83,36.54,37.23,37.96,38.67,39.39,40.17,40.93,41.74,42.5,43.28,44,44.75,45.47,46.2,46.91,47.6,48.28,48.97,49.62,50.28,50.89,51.51,52.11,52.72,53.28,53.89,54.45,55.05,55.62,56.19,56.77],
    p15: [14.58,14.84,15.1,15.37,15.64,15.9,16.21,16.47,16.77,17.1,17.44,17.77,18.11,18.47,18.8,19.18,19.57,19.98,20.37,20.78,21.21,21.65,22.08,22.52,22.98,23.44,23.91,24.41,24.88,25.35,25.85,26.45,27.09,27.72,28.36,28.95,29.58,30.22,30.82,31.46,32.09,32.8,33.52,34.27,35.03,35.82,36.66,37.49,38.37,39.19,40.04,40.86,41.74,42.57,43.41,44.24,45.1,45.97,46.9,47.85,48.8,49.82,50.84,51.89,52.95,54.03,55.08,56.14,57.26,58.41,59.63,60.85,62.1,63.32,64.57,65.76,66.98,68.13,69.29,70.47,71.63,72.78,73.9,75.01,76.07,77.15,78.14,79.14,80.15,81.11,82.12,83.08,84.02,84.98,85.96,86.9],
    p2: [15.2,15.47,15.74,16.04,16.3,16.6,16.91,17.2,17.51,17.84,18.21,18.51,18.91,19.27,19.67,20.08,20.47,20.88,21.31,21.75,22.21,22.68,23.14,23.62,24.11,24.61,25.12,25.61,26.15,26.61,27.15,27.85,28.56,29.25,29.96,30.66,31.35,32.03,32.72,33.43,34.12,34.93,35.73,36.54,37.4,38.26,39.17,40.09,41.04,41.93,42.87,43.8,44.71,45.64,46.55,47.44,48.37,49.31,50.3,51.32,52.34,53.42,54.51,55.63,56.75,57.9,59.02,60.18,61.36,62.59,63.9,65.19,66.57,67.9,69.21,70.5,71.78,73.04,74.3,75.58,76.8,78.05,79.24,80.46,81.58,82.73,83.82,84.88,85.96,86.98,88.03,89.05,90.1,91.09,92.14,93.14],
    p3: [16.68,16.97,17.27,17.58,17.87,18.2,18.54,18.87,19.24,19.61,19.98,20.37,20.81,21.21,21.64,22.11,22.54,23.05,23.51,24.02,24.51,25.05,25.58,26.15,26.72,27.28,27.85,28.41,29.02,29.62,30.13,30.99,31.86,32.76,33.63,34.49,35.49,36.46,37.42,38.43,39.39,40.37,41.4,42.41,43.47,44.56,45.71,46.83,48.01,49.1,50.25,51.37,52.49,53.55,54.66,55.72,56.78,57.89,59.01,60.16,61.35,62.6,63.82,65.1,66.33,67.64,68.9,70.22,71.61,73,74.48,76,77.59,79.11,80.63,82.11,83.6,85.05,86.51,87.96,89.38,90.8,92.16,93.51,94.83,96.11,97.33,98.52,99.74,100.9,102.04,103.2,104.35,105.44,106.59,107.7]
  },
  'หญิง': {
    m2: [10.65,10.83,11.02,11.2,11.38,11.57,11.77,11.95,12.13,12.34,12.52,12.73,12.94,13.17,13.39,13.61,13.85,14.09,14.33,14.6,14.87,15.14,15.41,15.7,15.98,16.26,16.56,16.87,17.19,17.49,17.79,18.09,18.41,18.72,18.99,19.22,19.44,19.68,19.91,20.15,20.42,20.67,20.92,21.2,21.46,21.74,22.04,22.34,22.61,22.9,23.2,23.52,23.89,24.28,24.67,25.06,25.52,25.97,26.47,27.01,27.59,28.21,28.88,29.63,30.4,31.14,31.92,32.68,33.39,34.11,34.79,35.44,36.07,36.66,37.22,37.72,38.23,38.68,39.15,39.63,40.09,40.56,41.02,41.47,41.92,42.36,42.78,43.18,43.56,43.95,44.33,44.7,45.07,45.43,45.76,46.12],
    m15: [11.1,11.3,11.48,11.69,11.88,12.07,12.27,12.46,12.66,12.88,13.09,13.29,13.5,13.73,13.97,14.21,14.46,14.72,14.97,15.27,15.54,15.81,16.1,16.4,16.71,17.03,17.34,17.66,17.96,18.29,18.62,18.95,19.28,19.61,19.89,20.14,20.41,20.68,20.95,21.19,21.48,21.78,22.08,22.36,22.65,22.95,23.28,23.61,23.91,24.23,24.57,24.93,25.32,25.74,26.19,26.65,27.16,27.67,28.2,28.8,29.44,30.14,30.9,31.71,32.53,33.35,34.17,34.98,35.75,36.48,37.2,37.9,38.56,39.18,39.76,40.3,40.84,41.35,41.83,42.33,42.86,43.38,43.88,44.39,44.87,45.34,45.82,46.27,46.72,47.15,47.59,48.02,48.43,48.82,49.24,49.63],
    p15: [14.51,14.77,15.01,15.28,15.54,15.8,16.05,16.32,16.59,16.89,17.16,17.46,17.76,18.08,18.4,18.73,19.07,19.43,19.79,20.18,20.57,20.96,21.37,21.79,22.21,22.64,23.09,23.54,23.99,24.47,24.92,25.39,25.85,26.32,26.81,27.31,27.82,28.34,28.87,29.38,29.89,30.42,30.96,31.48,32,32.56,33.2,33.82,34.42,35.04,35.68,36.36,37.11,37.9,38.69,39.53,40.41,41.34,42.29,43.3,44.39,45.53,46.73,47.96,49.19,50.39,51.53,52.61,53.62,54.59,55.51,56.4,57.27,58.1,58.91,59.71,60.49,61.27,62.07,62.92,63.81,64.71,65.61,66.55,67.47,68.39,69.36,70.29,71.18,72.08,73,73.91,74.81,75.71,76.61,77.52],
    p2: [15.21,15.47,15.74,16.01,16.28,16.57,16.85,17.13,17.43,17.73,18.03,18.33,18.65,18.98,19.33,19.68,20.05,20.41,20.8,21.22,21.61,22.05,22.48,22.92,23.38,23.84,24.3,24.78,25.26,25.77,26.25,26.73,27.23,27.74,28.27,28.84,29.41,29.98,30.54,31.11,31.69,32.27,32.87,33.46,34.04,34.65,35.35,36.07,36.75,37.44,38.16,38.93,39.74,40.58,41.46,42.36,43.33,44.32,45.35,46.46,47.64,48.87,50.14,51.44,52.74,53.98,55.21,56.32,57.33,58.33,59.3,60.23,61.13,62,62.86,63.69,64.54,65.41,66.27,67.18,68.15,69.14,70.14,71.19,72.21,73.25,74.3,75.32,76.36,77.38,78.4,79.42,80.44,81.49,82.51,83.53],
    p3: [16.87,17.17,17.47,17.77,18.07,18.37,18.68,19.01,19.34,19.67,20,20.35,20.71,21.09,21.48,21.87,22.27,22.7,23.14,23.59,24.04,24.54,25.03,25.53,26.05,26.57,27.09,27.64,28.2,28.74,29.31,29.87,30.44,30.96,31.54,32.22,32.92,33.64,34.32,35.01,35.81,36.58,37.36,38.14,38.93,39.77,40.7,41.63,42.55,43.48,44.44,45.39,46.36,47.38,48.43,49.49,50.63,51.77,52.99,54.24,55.58,56.93,58.35,59.8,61.2,62.55,63.84,65.04,66.11,67.14,68.16,69.16,70.13,71.09,72.05,73,73.99,74.99,76.04,77.16,78.33,79.54,80.79,82.06,83.36,84.69,86.01,87.34,88.64,89.98,91.31,92.64,93.98,95.34,96.72,98.06]
  }
};

function ensureCriteria_(ss) {
  var sh = ss.getSheetByName(SH.criteria);
  if (!sh) sh = ss.insertSheet(SH.criteria);
  var isNew = String(sh.getRange(1, 3).getValue()) === CRIT_HEADERS[2];
  if (!isNew) {
    if (sh.getLastRow() > 1) return sh; // มีข้อมูลรูปแบบอื่นอยู่ ไม่แตะต้อง
    sh.clear();
    sh.getRange(1, 1, 1, CRIT_HEADERS.length).setValues([CRIT_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() < 2) seedCriteria_(sh, CRIT_SEED, 65, '2-5');
  // เวอร์ชันก่อนหน้ามีเฉพาะ 2-5 ปี: เติมคอลัมน์ "ช่วงอายุ" และตาราง 6-19 ปี
  fixBandColumn_(sh);
  if (String(sh.getRange(1, 10).getValue()) !== 'ช่วงอายุ') {
    sh.getRange(1, 10).setValue('ช่วงอายุ').setFontWeight('bold');
    var last = sh.getLastRow();
    if (last > 1) {
      var fill = [];
      for (var k = 0; k < last - 1; k++) fill.push([BAND_LABEL['2-5']]);
      sh.getRange(2, 10, last - 1, 1).setNumberFormat('@');
      sh.getRange(2, 10, last - 1, 1).setValues(fill);
    }
    seedCriteria_(sh, CRIT_SEED_619, 90, '6-19');
  }
  return sh;
}

// Google Sheets แปลงข้อความอย่าง "2-5" เป็นวันที่เอง (2 พ.ค.) -> เก็บเป็นข้อความที่ไม่เหมือนวันที่ ("อายุ 2-5 ปี")
// และอ่านค่าเก่าที่ถูกแปลงแล้ว (วันที่, "5-2", "2-5", "6-19", "19-6") กลับเป็นช่วงอายุให้ถูก
var BAND_LABEL = { '2-5': 'อายุ 2-5 ปี', '6-19': 'อายุ 6-19 ปี' };
function bandText_(v) {
  var nums;
  if (v instanceof Date) {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    nums = Utilities.formatDate(v, tz, 'd-M').split('-').map(Number);
  } else {
    nums = (String(v == null ? '' : v).match(/\d+/g) || []).map(Number);
  }
  if (!nums.length) return '';
  return (nums.indexOf(6) >= 0 || nums.indexOf(19) >= 0) ? '6-19' : '2-5';
}
function fixBandColumn_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return;
  var rng = sh.getRange(2, 10, last - 1, 1), v = rng.getValues(), bad = false;
  var out = v.map(function (r) {
    var b = bandText_(r[0]);
    var lab = b ? BAND_LABEL[b] : '';
    if (String(r[0]) !== lab) bad = true;
    return [lab];
  });
  if (!bad) return;
  rng.setNumberFormat('@');
  rng.setValues(out);
}

// ต่อท้ายตารางเกณฑ์ (ไม่ทับของเดิม)
function seedCriteria_(sh, seed, h0, band) {
  var rows = [];
  ['ชาย', 'หญิง'].forEach(function (g) {
    var c = seed[g];
    for (var i = 0; i < c.m2.length; i++) {
      rows.push([g, h0 + i, '', c.m2[i], c.m15[i], '', c.p15[i], c.p2[i], c.p3[i], BAND_LABEL[band]]);
    }
  });
  var start = sh.getLastRow() + 1;
  sh.getRange(start, 10, rows.length, 1).setNumberFormat('@'); // กัน Sheets แปลง "2-5" เป็นวันที่
  sh.getRange(start, 1, rows.length, 10).setValues(rows);
}

function ensureSettings_(ss) {
  var sh = ss.getSheetByName(SH.settings);
  if (sh && String(sh.getRange(2, 3).getValue()) === 'key') { removeObsoleteSettings_(sh); return sh; }
  var old = {};
  if (sh) old = readOldSettings_(sh); else sh = ss.insertSheet(SH.settings);
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (e) {}
  sh.clear();
  var rows = SETTING_DEFS.map(function (d) {
    var v = (old[d[0]] !== undefined && old[d[0]] !== '') ? old[d[0]] : d[2];
    return [d[1], v, d[0]];
  });
  sh.getRange(1, 1).setValue('ตั้งค่าระบบ (แก้ไขได้จากหน้า "ตั้งค่า" ในแอป)');
  sh.getRange(2, 1, 1, 3).setValues([['รายการ', 'ค่า', 'key']]).setFontWeight('bold');
  sh.getRange(3, 1, rows.length, 3).setValues(rows);
  return sh;
}

// ลบแถวตั้งค่าที่เวอร์ชันนี้ไม่ใช้แล้ว (เช่น เกณฑ์ BMI)
function removeObsoleteSettings_(sh) {
  var last = sh.getLastRow();
  if (last < 3) return;
  var keys = {};
  SETTING_DEFS.forEach(function (d) { keys[d[0]] = 1; });
  var v = sh.getRange(3, 3, last - 2, 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    var k = String(v[i][0]);
    if (k && !keys[k]) sh.deleteRow(i + 3);
  }
}

// อ่านค่าจากชีต "ตั้งค่า" รูปแบบเก่า โดยจับคู่จากป้ายชื่อ (ไม่ผูกกับเลขแถว)
function readOldSettings_(sh) {
  var out = {};
  var last = Math.min(sh.getLastRow(), 40);
  if (last < 1) return out;
  var v = sh.getRange(1, 1, last, 2).getValues();
  var map = [
    ['ตำแหน่งครู', 'teacherPos'], ['ตำแหน่งผู้อำนวยการ', 'directorPos'],
    ['ชื่อ-สกุล ครู', 'teacherName'], ['ชื่อ-สกุล ผู้อำนวยการ', 'directorName'],
    ['ชื่อโรงเรียน', 'schoolName'], ['ระดับชั้น', 'grade'],
    ['ประจำปีการศึกษา', 'academicYear'], ['ประจำเดือน', 'currentMonth']
  ];
  v.forEach(function (r) {
    var label = String(r[0]).trim();
    var val = r[1];
    if (val === '' || val === null || String(val).charAt(0) === '#') return;
    for (var i = 0; i < map.length; i++) {
      if (label.indexOf(map[i][0]) === 0) {
        var key = map[i][1];
        if (key === 'currentMonth') {
          var idx = THAI_MONTHS.indexOf(String(val).trim());
          val = idx >= 0 ? idx + 1 : Number(val);
          if (!(val >= 1 && val <= 12)) return;
        }
        if (key === 'academicYear') { val = Number(val); if (!val) return; }
        out[key] = val;
        return;
      }
    }
  });
  return out;
}

function loadSettings_() {
  var sh = ensureSettings_(SpreadsheetApp.getActiveSpreadsheet());
  var last = sh.getLastRow();
  var vals = last >= 3 ? sh.getRange(3, 1, last - 2, 3).getValues() : [];
  var map = {};
  vals.forEach(function (r) { if (r[2]) map[String(r[2])] = r[1]; });
  var out = {};
  SETTING_DEFS.forEach(function (d) {
    var v = map[d[0]];
    if (typeof d[2] === 'number') {
      v = Number(v);
      if (v === '' || isNaN(v)) v = d[2];
    } else if (v === undefined || v === null || v === '') {
      v = d[2];
    }
    out[d[0]] = v;
  });
  return out;
}

function ensureRoster_(ss) {
  var sh = ss.getSheetByName(SH.roster);
  if (!sh) {
    sh = ss.insertSheet(SH.roster);
    sh.getRange(1, 1).setValue('ทะเบียนรายชื่อนักเรียน');
    sh.getRange(4, 1, 1, 8).setValues([['เลขที่', 'รหัสประจำตัว', 'ชื่อ - สกุล', 'ชื่อเล่น', 'เพศ', 'วัน/เดือน/ปีเกิด', 'หมายเหตุ', 'ห้อง']]).setFontWeight('bold');
  }
  return sh;
}

function rosterHeaderRow_(sh) {
  var n = Math.min(Math.max(sh.getLastRow(), 1), 15);
  var v = sh.getRange(1, 1, n, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === 'เลขที่') {
      if (String(sh.getRange(i + 1, 8).getValue()).trim() !== 'ห้อง') sh.getRange(i + 1, 8).setValue('ห้อง').setFontWeight('bold');
      return i + 1;
    }
  }
  // ไม่พบหัวตาราง: สร้างไว้ที่แถว 4
  sh.getRange(4, 1, 1, 8).setValues([['เลขที่', 'รหัสประจำตัว', 'ชื่อ - สกุล', 'ชื่อเล่น', 'เพศ', 'วัน/เดือน/ปีเกิด', 'หมายเหตุ', 'ห้อง']]).setFontWeight('bold');
  return 4;
}

function fmtDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v === null || v === undefined ? '' : v).trim();
}

function getRoster_(fresh) {
  var raw = readRosterRaw_(fresh);
  var cnt = {};
  return raw.map(function (s) {
    cnt[s.room] = (cnt[s.room] || 0) + 1;
    return { no: cnt[s.room], id: s.id, name: s.name, nickname: s.nickname, gender: s.gender,
             birth: s.birth, note: s.note, room: s.room };
  });
}

function ensureHolidays_(ss) {
  var sh = ss.getSheetByName(SH.holidays);
  if (!sh) {
    sh = ss.insertSheet(SH.holidays);
    sh.getRange(1, 2).setValue('ตารางวันหยุดนักขัตฤกษ์และวันหยุดพิเศษ (กรอกเอง หรือแก้จากหน้า "ตั้งค่า" ในแอป)');
    sh.getRange(3, 1, 1, 5).setValues([['', 'วันที่', 'วันในสัปดาห์', 'ชื่อวันหยุด', 'โรงเรียนเปิดสอนเป็นกรณีพิเศษ']]).setFontWeight('bold');
    sh.getRange(1, 2, sh.getMaxRows(), 1).setNumberFormat('@');
  }
  return sh;
}

function holidayHeaderRow_(sh) {
  var n = Math.min(Math.max(sh.getLastRow(), 1), 15);
  var v = sh.getRange(1, 2, n, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === 'วันที่') return i + 1;
  }
  return 3;
}

function loadHolidays_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureHolidays_(ss);
  var last = sh.getLastRow();
  if (last < 1) return [];
  var tz = ss.getSpreadsheetTimeZone();
  var out = [];
  sh.getRange(1, 1, last, 5).getValues().forEach(function (r) {
    var d = fmtDate_(r[1], tz);
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
      out.push({ date: d, name: String(r[3]).trim(), open: (r[4] === true || String(r[4]).toUpperCase() === 'TRUE') });
    }
  });
  return out;
}

function loadCriteria_() {
  var sh = ensureSetup_().criteria;
  var last = sh.getLastRow();
  var out = [];
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 10).getValues().forEach(function (r) {
    var o = { g: String(r[0]).trim(), b: bandText_(r[9]) || '2-5', h: parseFloat(r[1]), m2: parseFloat(r[3]),
              m15: parseFloat(r[4]), p15: parseFloat(r[6]), p2: parseFloat(r[7]), p3: parseFloat(r[8]) };
    // ต้องมีครบ 5 ค่าที่ใช้จัดกลุ่ม (-2, -1.5, +1.5, +2, +3) ส่วน -3SD และมัธยฐานเว้นว่างได้
    if (o.g && !isNaN(o.h) && !isNaN(o.m2) && !isNaN(o.m15) && !isNaN(o.p15) && !isNaN(o.p2) && !isNaN(o.p3)) {
      out.push(o);
    }
  });
  return out;
}

function loadYears_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var set = {};
  [SH.routine, SH.health].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last < 2) return;
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      var y = Number(r[0]);
      if (y) set[y] = 1;
    });
  });
  return Object.keys(set).map(Number).sort(function (a, b) { return b - a; });
}

function readRoutine_(year, month) {
  var sh = ensureSetup_().routine;
  var out = { milk: {}, brush: {} };
  var last = sh.getLastRow();
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 6).getValues().forEach(function (r) {
    if (Number(r[0]) === year && Number(r[1]) === month) {
      var t = String(r[2]);
      if (out[t]) {
        out[t][String(r[3])] = String(r[4]).split(',').filter(function (x) { return x !== ''; }).map(Number);
      }
    }
  });
  return out;
}

function readHealth_(year, month) {
  var sh = ensureSetup_().health;
  var out = {};
  var last = sh.getLastRow();
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 5).getValues().forEach(function (r) {
    if (Number(r[0]) === year && Number(r[1]) === month) {
      out[String(r[2])] = { w: r[3] === '' ? '' : Number(r[3]), h: r[4] === '' ? '' : Number(r[4]) };
    }
  });
  return out;
}

function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    var token = contents.token;
    var fn = contents.fn;
    var args = contents.args || [];
    
    var result = rpc(token, fn, args);
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: err.toString(), code: 'ERR' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}