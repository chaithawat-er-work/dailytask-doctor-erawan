/**
 * ระบบตารางการทำงานแพทย์ (Doctor Schedule) — Google Apps Script
 * ---------------------------------------------------------------
 * Part 1  : Web App ให้ผู้ใช้กรอก/แก้ไข/ลบรายการงาน  (doGet + API functions)
 * Part 2  : แสดงผล — ตารางรายสัปดาห์ใน Web App และ render ชีต "Calendar" แบบเดิม
 * Part 3  : Email notify — สรุปวันพรุ่งนี้ทุกเย็น + สรุปสัปดาห์หน้าเช้าวันอาทิตย์
 *
 * วิธีติดตั้ง: ดู README.md
 */

// ====================== CONFIG ======================
const SHEETS = {
  STAFF: 'Staff',        // รายชื่อแพทย์ + email
  SCHEDULE: 'Schedule',  // 1 แถว = 1 รายการงาน (source of truth)
  TRAINEE: 'Trainee',    // Fellow / Resident / ENP / Paramedic ต่อวัน
  CONFIG: 'Config',      // ตั้งค่าระบบ
  CALENDAR: 'Calendar',  // ปฏิทินที่ render อัตโนมัติ (ห้ามแก้มือ)
};

const SHIFTS = ['เช้า', 'บ่าย', 'อื่นๆ', 'ทั้งวัน'];
const TRAINEE_TYPES = ['Fellow', 'Resident', 'ENP', 'Paramedic'];
const LOCATIONS = ['ในศูนย์เอราวัณ', 'นอกสถานที่'];
const MEETING_LOCATIONS = Object.freeze({
  'ประชุมในเอราวัณประจำ': LOCATIONS[0],
  'ประชุมในเอราวัณไม่ประจำ': LOCATIONS[0],
  'ประชุมนอกเอราวัณประจำ': LOCATIONS[1],
  'ประชุมนอกเอราวัณไม่ประจำ': LOCATIONS[1],
  'ประชุมออนไลน์ประจำ': 'ออนไลน์',
});
const TZ = 'Asia/Bangkok';
const THAI_DAYS = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
const THAI_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

// Location ต่อท้ายสุด เพื่อไม่ให้กระทบข้อมูลเดิมที่มีอยู่แล้ว
const SCHEDULE_HEADER = ['ID', 'Date', 'Name', 'Shift', 'Activity', 'Note', 'RecordedBy', 'Timestamp', 'Location'];
const MEETING_HEADER = 'MeetingType';
const TRAINEE_HEADER = ['ID', 'Date', 'Type', 'Institution', 'Count', 'Note', 'RecordedBy', 'Timestamp'];
const STAFF_HEADER = ['Name', 'Role', 'Email', 'Active', 'Order'];
const CONFIG_HEADER = ['Key', 'Value', 'Description'];

// ====================== SETUP ======================
/** รันครั้งแรกครั้งเดียว: สร้างชีตทั้งหมด + ข้อมูลตัวอย่าง */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const staff = getOrCreateSheet_(ss, SHEETS.STAFF, STAFF_HEADER);
  if (staff.getLastRow() === 1) {
    staff.getRange(2, 1, 5, 5).setValues([
      ['แพทย์ตัวอย่าง 1', 'แพทย์', '', true, 1],
      ['แพทย์ตัวอย่าง 2', 'แพทย์', '', true, 2],
      ['แพทย์ตัวอย่าง 3', 'แพทย์', '', true, 3],
      ['บุคลากรตัวอย่าง', 'แพทย์', '', true, 4],
      ['ผอ.ส่วนบริการ (ตัวอย่าง)', 'ผอส่วน', '', true, 9],
    ]);
  }

  const sched = getOrCreateSheet_(ss, SHEETS.SCHEDULE, SCHEDULE_HEADER);
  ensureMeetingColumn_(sched);
  applyScheduleValidation(sched);
  getOrCreateSheet_(ss, SHEETS.TRAINEE, TRAINEE_HEADER);

  const cfg = getOrCreateSheet_(ss, SHEETS.CONFIG, CONFIG_HEADER);
  if (cfg.getLastRow() === 1) {
    cfg.getRange(2, 1, 6, 3).setValues([
      ['ADMIN_EMAILS', Session.getEffectiveUser().getEmail(), 'อีเมลผู้รับสรุป (คั่นด้วย , )'],
      ['DAILY_HOUR', 17, 'ชั่วโมงที่ส่งสรุปวันพรุ่งนี้ (0-23)'],
      ['WEEKLY_HOUR', 8, 'ชั่วโมงที่ส่งสรุปสัปดาห์ (วันอาทิตย์)'],
      ['SEND_TO_STAFF', true, 'ส่งเมลถึงแพทย์แต่ละคนด้วยหรือไม่ (ใช้ Email ในชีต Staff)'],
      ['ORG_NAME', 'ศูนย์เอราวัณ', 'ชื่อหน่วยงานในหัวเมล'],
      ['WEBAPP_URL', '', 'URL ของ Web App (ใส่หลัง Deploy เพื่อแนบลิงก์ในเมล)'],
    ]);
  }

  getOrCreateSheet_(ss, SHEETS.CALENDAR, null);
  renderCalendar();
  SpreadsheetApp.getUi().alert('ติดตั้งเสร็จแล้ว ✅\nถัดไป: Deploy > New deployment > Web app แล้วรัน setupTriggers()');
}

/** ติดตั้ง trigger สำหรับ email + auto render */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  const dailyHour = Number(getConfig_('DAILY_HOUR', 17));
  const weeklyHour = Number(getConfig_('WEEKLY_HOUR', 8));

  ScriptApp.newTrigger('sendTomorrowDigest').timeBased().everyDays(1).atHour(dailyHour).inTimezone(TZ).create();
  ScriptApp.newTrigger('sendWeeklyDigest').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(weeklyHour).inTimezone(TZ).create();
  ScriptApp.newTrigger('renderCalendar').timeBased().everyHours(1).create();
  Logger.log('Triggers installed: daily %s:00, weekly Sunday %s:00, render hourly', dailyHour, weeklyHour);
}

/** dropdown ในชีต Schedule สำหรับ admin ที่แก้ตรง (Name / Shift / Location) */
function applyScheduleValidation(sh) {
  sh = sh || sheet_(SHEETS.SCHEDULE);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = h => hdr.indexOf(h) + 1;
  const n = sh.getMaxRows() - 1;
  const staffSh = sheet_(SHEETS.STAFF);
  sh.getRange(2, col('Name'), n).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(staffSh.getRange('A2:A'), true).setAllowInvalid(false).build());
  sh.getRange(2, col('Shift'), n).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(SHIFTS, true).setAllowInvalid(false).build());
  applyScheduleLocationValidation_(sh, col('Location'), col(MEETING_HEADER));
}

function applyScheduleLocationValidation_(sh, locationCol, meetingCol) {
  const n = sh.getMaxRows() - 1;
  sh.getRange(2, locationCol, n).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(meetingCol ? LOCATIONS.concat('ออนไลน์') : LOCATIONS, true).setAllowInvalid(false).build());
  if (meetingCol) {
    sh.getRange(2, meetingCol, n).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.keys(MEETING_LOCATIONS), true).setAllowInvalid(false).build());
  }
}

/** อัปเกรดชีต Schedule เดิมจากตัวแก้ไข Apps Script โดยไม่เรียก setup หรือสร้าง trigger ใหม่ */
function setupMeetingTypes() {
  requireAdmin_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(SHEETS.SCHEDULE);
    ensureMeetingColumn_(sh);
    applyScheduleLocationValidation_(sh, SCHEDULE_HEADER.length, SCHEDULE_HEADER.length + 1);
  } finally {
    lock.releaseLock();
  }
}

function scheduleHasMeetingColumn_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (SCHEDULE_HEADER.some((header, i) => headers[i] !== header)) {
    throw new Error('หัวคอลัมน์ชีต Schedule ไม่ตรงกับโครงสร้างที่ระบบรองรับ');
  }
  const meetingIndex = headers.indexOf(MEETING_HEADER);
  if (meetingIndex >= 0 && meetingIndex !== SCHEDULE_HEADER.length) {
    throw new Error('คอลัมน์ MeetingType ต้องอยู่ถัดจาก Location');
  }
  return meetingIndex >= 0;
}

function ensureMeetingColumn_(sh) {
  if (scheduleHasMeetingColumn_(sh)) return;
  if (sh.getLastColumn() !== SCHEDULE_HEADER.length) {
    throw new Error('พบคอลัมน์เพิ่มเติมในชีต Schedule กรุณาตรวจโครงสร้างก่อนเพิ่ม MeetingType');
  }
  const column = SCHEDULE_HEADER.length + 1;
  if (sh.getMaxColumns() < column) sh.insertColumnAfter(column - 1);
  sh.getRange(1, column).setValue(MEETING_HEADER).setFontWeight('bold').setBackground('#dbe9f6');
}

/**
 * สิทธิ์ admin: เจ้าของสคริปต์ หรืออีเมลที่อยู่ใน Config › ADMIN_EMAILS
 * ใช้กับทุกฟังก์ชันที่แก้ไขรายชื่อ staff
 */
function isAdmin_() {
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const owner = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!me) return false;
  if (me === owner) return true;
  return String(getConfig_('ADMIN_EMAILS', '')).toLowerCase().split(',').map(x => x.trim()).filter(Boolean).indexOf(me) >= 0;
}

function requireAdmin_() {
  if (!isAdmin_()) throw new Error('เฉพาะผู้ดูแลระบบเท่านั้นที่แก้ไขรายชื่อได้ (เพิ่มอีเมลของคุณใน Config › ADMIN_EMAILS)');
}

/** ล็อกชีต Staff ให้แก้ได้เฉพาะ admin (คนอื่นยังเปิดดูได้) */
function protectStaffSheet() {
  requireAdmin_();
  const sh = sheet_(SHEETS.STAFF);
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => { if (p.canEdit()) p.remove(); });
  const p = sh.protect().setDescription('รายชื่อ staff — แก้ไขได้เฉพาะผู้ดูแลระบบ');
  p.removeEditors(p.getEditors().map(u => u.getEmail()));
  const admins = String(getConfig_('ADMIN_EMAILS', '')).split(',').map(x => x.trim()).filter(Boolean);
  if (admins.length) p.addEditors(admins);
  if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false);
  SpreadsheetApp.getUi().alert('ล็อกชีต Staff แล้ว ✅\nแก้ไขได้เฉพาะ: ' + [Session.getEffectiveUser().getEmail()].concat(admins).join(', '));
}

/** เมนู admin: แก้ไขชื่อ staff (อัปเดตรายการงานเดิมให้ด้วย) */
function renameStaffPrompt() {
  requireAdmin_();
  const ui = SpreadsheetApp.getUi();
  const list = getStaff_();
  const r1 = ui.prompt('แก้ไขชื่อ', 'ชื่อเดิม (พิมพ์ให้ตรง)\n\nรายชื่อปัจจุบัน: ' + list.map(s => s.Name).join(', '), ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const oldName = r1.getResponseText().trim();
  const target = list.find(s => s.Name === oldName);
  if (!target) { ui.alert('ไม่พบชื่อ "' + oldName + '"'); return; }
  const r2 = ui.prompt('แก้ไขชื่อ', 'ชื่อใหม่ของ "' + oldName + '"', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK || !r2.getResponseText().trim()) return;
  const newName = r2.getResponseText().trim();
  if (list.some(s => s.Name === newName)) { ui.alert('มีชื่อ "' + newName + '" อยู่แล้ว'); return; }

  sheet_(SHEETS.STAFF).getRange(target._row, 1).setValue(newName);
  // อัปเดตชื่อในรายการงานเดิมทั้งหมด
  const sch = sheet_(SHEETS.SCHEDULE);
  const hdr = sch.getRange(1, 1, 1, sch.getLastColumn()).getValues()[0];
  const nameCol = hdr.indexOf('Name') + 1;
  const n = sch.getLastRow() - 1;
  let changed = 0;
  if (n > 0 && nameCol > 0) {
    const rng = sch.getRange(2, nameCol, n, 1);
    const vals = rng.getValues();
    vals.forEach(r => { if (r[0] === oldName) { r[0] = newName; changed++; } });
    rng.setValues(vals);
  }
  applyScheduleValidation(sch);
  renderCalendar();
  ui.alert('เปลี่ยนชื่อเป็น "' + newName + '" แล้ว ✅\nอัปเดตรายการงานเดิม ' + changed + ' รายการ');
}

/** เมนู admin: เพิ่มแพทย์ใหม่ */
function addStaffPrompt() {
  requireAdmin_();
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('เพิ่มแพทย์ใหม่', 'ชื่อที่จะแสดงในตาราง (เช่น สมชาย)', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK || !r1.getResponseText().trim()) return;
  const name = r1.getResponseText().trim();
  if (getStaff_().some(s => s.Name === name)) { ui.alert('มีชื่อ "' + name + '" อยู่แล้ว'); return; }
  const r2 = ui.prompt('อีเมลของ ' + name, 'สำหรับรับเมลสรุป (เว้นว่างได้)', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const sh = sheet_(SHEETS.STAFF);
  const maxOrder = getStaff_().filter(s => s.Role !== 'ผอส่วน').reduce((m, s) => Math.max(m, Number(s.Order) || 0), 0);
  sh.appendRow([name, 'แพทย์', r2.getResponseText().trim(), true, maxOrder + 1]);
  renderCalendar();
  ui.alert('เพิ่ม ' + name + ' แล้ว ✅\nชื่อจะขึ้นใน Web App, ชีต Calendar และเมลสรุปทันที');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📅 ตารางแพทย์')
    .addItem('เปิดหน้ากรอกข้อมูล (Sidebar)', 'showSidebar')
    .addItem('Render ปฏิทินใหม่', 'renderCalendar')
    .addItem('➕ เพิ่มแพทย์ใหม่ (admin)', 'addStaffPrompt')
    .addItem('✏️ แก้ไขชื่อแพทย์ (admin)', 'renameStaffPrompt')
    .addItem('🔒 ล็อกชีต Staff ให้เฉพาะ admin', 'protectStaffSheet')
    .addItem('อัปเดต dropdown ชีต Schedule', 'applyScheduleValidation')
    .addSeparator()
    .addItem('ทดสอบส่งเมลสรุปพรุ่งนี้', 'sendTomorrowDigest')
    .addItem('ทดสอบส่งเมลสรุปสัปดาห์', 'sendWeeklyDigest')
    .addSeparator()
    .addItem('ติดตั้งระบบ (setup)', 'setup')
    .addItem('ติดตั้ง Trigger', 'setupTriggers')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Index').evaluate().setTitle('ตารางแพทย์');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ====================== PART 1: WEB APP ======================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ตารางการทำงานแพทย์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** ข้อมูลเริ่มต้นที่หน้าเว็บต้องใช้ */
function apiGetMeta() {
  return {
    staff: getStaff_().map(s => ({ name: s.Name, role: s.Role, hasEmail: !!String(s.Email || '').trim() })),
    shifts: SHIFTS,
    traineeTypes: TRAINEE_TYPES,
    locations: LOCATIONS,
    meetingLocations: MEETING_LOCATIONS,
    meetingTypesEnabled: scheduleHasMeetingColumn_(sheet_(SHEETS.SCHEDULE)),
    user: Session.getActiveUser().getEmail() || 'anonymous',
    isAdmin: isAdmin_(),
    today: fmtDate_(new Date()),
  };
}

/** บันทึกรายการงาน (ถ้ามี id = แก้ไข) */
function apiSaveSchedule(item) {
  validate_(item, ['date', 'name', 'shift']);
  if (SHIFTS.indexOf(item.shift) < 0) throw new Error('Shift ไม่ถูกต้อง');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(SHEETS.SCHEDULE);
    const hasMeetingColumn = scheduleHasMeetingColumn_(sh);
    const r = item.id ? findRowById_(sh, item.id) : -1;
    const meetingType = item.meetingType === undefined && hasMeetingColumn && r > 0
      ? sh.getRange(r, SCHEDULE_HEADER.length + 1).getValue()
      : item.meetingType;
    item = normalizeScheduleMeeting_({ ...item, meetingType }, hasMeetingColumn);
    const user = Session.getActiveUser().getEmail() || 'anonymous';
    const row = scheduleRow_(item, item.date, item.name.trim(), user, new Date(), hasMeetingColumn, item.id);
    if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return { ok: true, id: row[0] };
  } finally {
    lock.releaseLock();
  }
}

/** บันทึกหลายวันพร้อมกัน (เช่น ลาพักผ่อน 3 วัน) */
function apiSaveScheduleRange(item) {
  validate_(item, ['dateFrom', 'dateTo', 'name', 'shift']);
  const from = parseDate_(item.dateFrom), to = parseDate_(item.dateTo);
  if (to < from) throw new Error('วันสิ้นสุดต้องไม่ก่อนวันเริ่ม');
  if ((to - from) / 864e5 > 62) throw new Error('ช่วงวันเกิน 62 วัน');
  if (SHIFTS.indexOf(item.shift) < 0) throw new Error('Shift ไม่ถูกต้อง');
  const sh = sheet_(SHEETS.SCHEDULE);
  const hasMeetingColumn = scheduleHasMeetingColumn_(sh);
  item = normalizeScheduleMeeting_(item, hasMeetingColumn);
  const user = Session.getActiveUser().getEmail() || 'anonymous';
  const now = new Date();
  const rows = dateList_(from, to).map(dt => scheduleRow_(item, dt, item.name.trim(), user, now, hasMeetingColumn));
  appendRows_(sh, rows);
  return { ok: true, count: rows.length };
}

/**
 * บันทึกงานเดียวให้แพทย์หลายคน (และหลายวัน) ในครั้งเดียว
 * item: { names:[], dateFrom, dateTo, shift, activity, note, location, meetingType, notify }
 */
function apiSaveScheduleBulk(item) {
  validate_(item, ['dateFrom', 'shift', 'activity']);
  const names = (item.names || []).map(n => String(n).trim()).filter(Boolean);
  if (!names.length) throw new Error('กรุณาเลือกแพทย์อย่างน้อย 1 คน');
  if (SHIFTS.indexOf(item.shift) < 0) throw new Error('Shift ไม่ถูกต้อง');
  const staff = getStaff_();
  const unknown = names.filter(n => !staff.some(s => s.Name === n));
  if (unknown.length) throw new Error('ไม่พบชื่อในชีต Staff: ' + unknown.join(', '));

  const from = parseDate_(item.dateFrom), to = parseDate_(item.dateTo || item.dateFrom);
  if (to < from) throw new Error('วันสิ้นสุดต้องไม่ก่อนวันเริ่ม');
  if ((to - from) / 864e5 > 62) throw new Error('ช่วงวันเกิน 62 วัน');

  const sh = sheet_(SHEETS.SCHEDULE);
  const hasMeetingColumn = scheduleHasMeetingColumn_(sh);
  item = normalizeScheduleMeeting_(item, hasMeetingColumn);
  const user = Session.getActiveUser().getEmail() || 'anonymous';
  const now = new Date();
  const dates = dateList_(from, to);
  const rows = [];
  dates.forEach(dt => names.forEach(n => rows.push(scheduleRow_(item, dt, n, user, now, hasMeetingColumn))));
  appendRows_(sh, rows);

  let mailed = 0;
  if (item.notify) mailed = notifyAssignees_(names, dates, item, staff);
  return { ok: true, count: rows.length, days: dates.length, people: names.length, mailed: mailed };
}

function normalizeScheduleMeeting_(item, hasMeetingColumn) {
  const meetingType = item.meetingType || '';
  if (meetingType && (typeof meetingType !== 'string' || !Object.prototype.hasOwnProperty.call(MEETING_LOCATIONS, meetingType))) {
    throw new Error('ประเภทการประชุมไม่ถูกต้อง');
  }
  if (meetingType && !hasMeetingColumn) throw new Error('กรุณาให้ผู้ดูแลเปิดใช้ประเภทการประชุมในชีต Schedule ก่อน');
  const location = meetingType ? MEETING_LOCATIONS[meetingType] : (item.location || LOCATIONS[0]);
  if (!meetingType && !LOCATIONS.includes(location)) throw new Error('สถานที่ไม่ถูกต้อง');
  return { ...item, meetingType, location };
}

function scheduleRow_(item, date, name, user, now, hasMeetingColumn, id) {
  const row = [id || Utilities.getUuid(), date, name, item.shift, (item.activity || '').trim(),
    (item.note || '').trim(), user, now, item.location];
  if (hasMeetingColumn) row.push(item.meetingType);
  return row;
}

/** ส่งเมลแจ้งแพทย์ที่ถูกมอบหมายงาน (เฉพาะคนที่มี Email ในชีต Staff) */
function notifyAssignees_(names, dates, item, staff) {
  staff = staff || getStaff_();
  const to = names.map(n => (staff.find(s => s.Name === n) || {}).Email).filter(Boolean).map(e => String(e).trim());
  if (!to.length) return 0;
  const when = dates.length > 1 ? thaiDate_(dates[0]) + ' – ' + thaiDate_(dates[dates.length - 1]) : thaiDate_(dates[0]);
  const org = getConfig_('ORG_NAME', '');
  const url = getConfig_('WEBAPP_URL', '');
  const html = `<div style="font-family:Arial,'Helvetica Neue',sans-serif;font-size:14px;color:#222">
    <h2 style="margin:0 0 10px">มีงานใหม่ในตาราง</h2>
    <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#ccc">
      <tr><td><b>กิจกรรม</b></td><td>${esc_(item.activity)}</td></tr>
      ${item.meetingType ? `<tr><td><b>ประเภทการประชุม</b></td><td>${esc_(item.meetingType)}</td></tr>` : ''}
      <tr><td><b>วันที่</b></td><td>${esc_(when)}${dates.length > 1 ? ' (' + dates.length + ' วัน)' : ''}</td></tr>
      <tr><td><b>ช่วงเวลา</b></td><td>${esc_(item.shift)}</td></tr>
      <tr><td><b>สถานที่</b></td><td>${item.location === 'นอกสถานที่' ? '🚗 ' : ''}${esc_(item.location || LOCATIONS[0])}</td></tr>
      <tr><td><b>ผู้เข้าร่วม</b></td><td>${esc_(names.join(', '))}</td></tr>
      ${item.note ? `<tr><td><b>หมายเหตุ</b></td><td>${esc_(item.note)}</td></tr>` : ''}
    </table>
    ${url ? `<p style="margin-top:14px"><a href="${url}" style="background:#1a73e8;color:#fff;padding:8px 14px;border-radius:4px;text-decoration:none">เปิดตาราง</a></p>` : ''}
    <p style="color:#999;font-size:12px">บันทึกโดย ${esc_(Session.getActiveUser().getEmail() || 'ระบบ')}</p></div>`;
  MailApp.sendEmail({ to: to.join(','), subject: `[${org}] งานใหม่: ${item.activity} · ${when}`, htmlBody: html, name: org || 'ตารางแพทย์' });
  return to.length;
}

/** บันทึก trainee หลายวันในครั้งเดียว (เขียนชีตรอบเดียว เร็ว) */
function apiSaveTraineeRange(item) {
  validate_(item, ['dateFrom', 'dateTo', 'type']);
  if (TRAINEE_TYPES.indexOf(item.type) < 0) throw new Error('ประเภทไม่ถูกต้อง');
  const from = parseDate_(item.dateFrom), to = parseDate_(item.dateTo);
  if (to < from) throw new Error('วันสิ้นสุดต้องไม่ก่อนวันเริ่ม');
  if ((to - from) / 864e5 > 62) throw new Error('ช่วงวันเกิน 62 วัน');
  const user = Session.getActiveUser().getEmail() || 'anonymous';
  const now = new Date();
  const rows = dateList_(from, to).map(dt => [Utilities.getUuid(), dt, item.type, (item.institution || '').trim(),
    Number(item.count) || 0, (item.note || '').trim(), user, now]);
  appendRows_(sheet_(SHEETS.TRAINEE), rows);
  return { ok: true, count: rows.length };
}

function dateList_(from, to) {
  const out = [];
  for (let d = new Date(from.getFullYear(), from.getMonth(), from.getDate()); d <= to; d.setDate(d.getDate() + 1)) {
    out.push(Utilities.formatDate(d, TZ, 'yyyy-MM-dd'));
  }
  return out;
}

function appendRows_(sh, rows) {
  if (!rows.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
}

function apiDeleteSchedule(id) {
  const sh = sheet_(SHEETS.SCHEDULE);
  const r = findRowById_(sh, id);
  if (r > 0) sh.deleteRow(r);
  return { ok: r > 0 };
}

function apiSaveTrainee(item) {
  validate_(item, ['date', 'type']);
  if (TRAINEE_TYPES.indexOf(item.type) < 0) throw new Error('ประเภทไม่ถูกต้อง');
  const sh = sheet_(SHEETS.TRAINEE);
  const user = Session.getActiveUser().getEmail() || 'anonymous';
  const row = [item.id || Utilities.getUuid(), item.date, item.type, (item.institution || '').trim(),
    Number(item.count) || 0, (item.note || '').trim(), user, new Date()];
  const r = item.id ? findRowById_(sh, item.id) : -1;
  if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: row[0] };
}

function apiDeleteTrainee(id) {
  const sh = sheet_(SHEETS.TRAINEE);
  const r = findRowById_(sh, id);
  if (r > 0) sh.deleteRow(r);
  return { ok: r > 0 };
}

// ====================== PART 2: DISPLAY ======================
/** ข้อมูลสัปดาห์ (จันทร์–ศุกร์ หรือ 7 วัน) สำหรับ Web App */
function apiGetWeek(anyDateStr, days) {
  const start = weekStart_(parseDate_(anyDateStr || fmtDate_(new Date())));
  const n = Number(days) || 5;
  const dates = [];
  for (let i = 0; i < n; i++) {
    const d = new Date(start); d.setDate(d.getDate() + i);
    dates.push(fmtDate_(d));
  }
  return { start: dates[0], end: dates[dates.length - 1], days: buildDays_(dates) };
}

/** ข้อมูลทั้งเดือน สำหรับมุมมองรายเดือน (ส่งทุกวันของเดือน + flag วันหยุด) */
function apiGetMonth(anyDateStr) {
  const d0 = parseDate_(anyDateStr || fmtDate_(new Date()));
  const first = new Date(d0.getFullYear(), d0.getMonth(), 1);
  const last = new Date(d0.getFullYear(), d0.getMonth() + 1, 0);
  const dates = dateList_(first, last);
  return {
    start: dates[0], end: dates[dates.length - 1],
    title: THAI_MONTHS_FULL[first.getMonth()] + ' ' + (first.getFullYear() + 543),
    staff: getStaff_().map(s => ({ name: s.Name, role: s.Role, hasEmail: !!String(s.Email || '').trim() })),
    days: buildDays_(dates),
  };
}

const THAI_MONTHS_FULL = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const THAI_DAYS_SHORT = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];

function buildDays_(dates) {
  const from = dates[0], to = dates[dates.length - 1];
  const staff = getStaff_();
  const schedule = getSchedule_(from, to);
  const trainee = getTrainee_(from, to);

  return dates.map(dt => {
    const d = parseDate_(dt);
    return {
      date: dt,
      label: THAI_DAYS[d.getDay()] + ' ' + d.getDate() + ' ' + THAI_MONTHS[d.getMonth()],
      dow: d.getDay(), dayNum: d.getDate(), dowShort: THAI_DAYS_SHORT[d.getDay()],
      staff: staff.map(s => {
        const cell = {};
        SHIFTS.forEach(sh => cell[sh] = []);
        schedule.filter(x => x.Date === dt && x.Name === s.Name).forEach(x => (cell[x.Shift] || cell['อื่นๆ']).push(x));
        return { name: s.Name, role: s.Role, shifts: cell };
      }),
      trainee: TRAINEE_TYPES.map(t => ({ type: t, items: trainee.filter(x => x.Date === dt && x.Type === t) })),
    };
  });
}

/** Render ชีต Calendar ให้หน้าตาเหมือนตารางเดิม (สัปดาห์ละ block, จันทร์–ศุกร์) */
function renderCalendar(weeksBack, weeksForward) {
  // trigger/เมนูจะส่ง event object มา → ใช้ค่า default
  weeksBack = typeof weeksBack === 'number' ? weeksBack : 1;
  weeksForward = typeof weeksForward === 'number' ? weeksForward : 5;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.CALENDAR, null);
  sh.clear(); sh.clearFormats();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();

  const staff = getStaff_();
  const doctors = staff.filter(s => s.Role !== 'ผอส่วน');
  const director = staff.filter(s => s.Role === 'ผอส่วน');
  const monday = weekStart_(new Date());
  monday.setDate(monday.getDate() - weeksBack * 7);
  const totalWeeks = weeksBack + weeksForward + 1;

  const from = fmtDate_(monday);
  const endD = new Date(monday); endD.setDate(endD.getDate() + totalWeeks * 7);
  const to = fmtDate_(endD);
  const schedule = getSchedule_(from, to);
  const trainee = getTrainee_(from, to);

  const COLS_PER_DAY = 5; // วันที่ | ชื่อ | เช้า | บ่าย | อื่นๆ
  const rowsPerBlock = 2 + doctors.length + TRAINEE_TYPES.length + director.length + 1; // header + subheader + ... + spacer
  const out = [];
  const merges = [];
  const styles = { bold: [], dayHead: [], weekend: [], trainee: [] };

  let r = 1;
  sh.getRange(1, 1).setValue('ตารางการทำงานแพทย์ (auto-generated ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') + ') — แก้ไขข้อมูลผ่าน Web App หรือชีต Schedule เท่านั้น');
  r = 3;

  for (let w = 0; w < totalWeeks; w++) {
    const wkStart = new Date(monday); wkStart.setDate(wkStart.getDate() + w * 7);
    const block = [];
    for (let i = 0; i < rowsPerBlock; i++) block.push(new Array(COLS_PER_DAY * 5).fill(''));

    for (let d = 0; d < 5; d++) {
      const day = new Date(wkStart); day.setDate(day.getDate() + d);
      const dt = fmtDate_(day);
      const c0 = d * COLS_PER_DAY;
      block[0][c0] = day.getDate() + ' ' + THAI_MONTHS[day.getMonth()];
      block[0][c0 + 1] = THAI_DAYS[day.getDay()];
      block[1][c0 + 2] = 'เช้า'; block[1][c0 + 3] = 'บ่าย'; block[1][c0 + 4] = 'อื่นๆ';
      styles.dayHead.push([r, c0 + 1, 2, COLS_PER_DAY]);

      doctors.forEach((s, i) => {
        const rr = 2 + i;
        block[rr][c0 + 1] = s.Name;
        const items = schedule.filter(x => x.Date === dt && x.Name === s.Name);
        const txt = sh => items.filter(x => x.Shift === sh || x.Shift === 'ทั้งวัน').map(label_).join(' / ');
        styles.offsite = styles.offsite || [];
        ['เช้า', 'บ่าย', 'อื่นๆ'].forEach((k, j) => { if (items.some(x => (x.Shift === k || x.Shift === 'ทั้งวัน') && isOffsite_(x))) styles.offsite.push([r + rr, c0 + 3 + j]); });
        block[rr][c0 + 2] = txt('เช้า');
        block[rr][c0 + 3] = txt('บ่าย');
        block[rr][c0 + 4] = txt('อื่นๆ');
      });
      let rr = 2 + doctors.length;
      TRAINEE_TYPES.forEach((t, i) => {
        const items = trainee.filter(x => x.Date === dt && x.Type === t);
        block[rr + i][c0 + 1] = i === 0 ? 'แพทย์ elective' : '';
        block[rr + i][c0 + 2] = t;
        block[rr + i][c0 + 3] = items.map(x => x.Institution).filter(Boolean).join(' + ');
        block[rr + i][c0 + 4] = items.reduce((a, x) => a + (Number(x.Count) || 0), 0) || '';
      });
      rr += TRAINEE_TYPES.length;
      director.forEach((s, i) => {
        block[rr + i][c0 + 1] = s.Name;
        const items = schedule.filter(x => x.Date === dt && x.Name === s.Name);
        block[rr + i][c0 + 2] = items.map(x => (x.Shift === 'ทั้งวัน' ? '' : x.Shift + ': ') + label_(x)).join(' / ');
        merges.push([r + rr + i, c0 + 3, 1, 3]);
      });
      // merge วันที่ลงแนวตั้งทั้ง block
      merges.push([r, c0 + 1, rowsPerBlock - 1, 1]);
      styles.trainee.push([r + 2 + doctors.length, c0 + 1, TRAINEE_TYPES.length, COLS_PER_DAY]);
    }
    out.push(...block);
    r += rowsPerBlock;
  }

  const range = sh.getRange(3, 1, out.length, COLS_PER_DAY * 5);
  range.setValues(out).setVerticalAlignment('middle').setWrap(true).setFontSize(10);
  range.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  styles.dayHead.forEach(a => sh.getRange(a[0], a[1], a[2], a[3]).setBackground('#dbe9f6').setFontWeight('bold').setHorizontalAlignment('center'));
  styles.trainee.forEach(a => sh.getRange(a[0], a[1], a[2], a[3]).setBackground('#f5f5f5').setFontColor('#555555'));
  (styles.offsite || []).forEach(a => sh.getRange(a[0], a[1]).setBackground('#fde8d7'));
  merges.forEach(a => sh.getRange(a[0], a[1], a[2], a[3]).merge());
  const legendRow = 3 + out.length + 1;
  sh.getRange(legendRow, 1, 1, 3).merge().setValue('🚗 = นอกสถานที่ (พื้นสีส้ม)   ·   ไม่มีเครื่องหมาย = ในศูนย์เอราวัณ').setFontColor('#555555');
  // ทำสีวันนี้
  const todayStr = fmtDate_(new Date());
  const offset = Math.round((parseDate_(todayStr) - monday) / 864e5);
  if (offset >= 0) {
    const wk = Math.floor(offset / 7), dow = offset % 7;
    if (dow < 5) sh.getRange(3 + wk * rowsPerBlock, dow * COLS_PER_DAY + 1, rowsPerBlock - 1, COLS_PER_DAY).setBackground('#fff8dc');
  }
  for (let d = 0; d < 5; d++) {
    sh.setColumnWidth(d * COLS_PER_DAY + 1, 60);
    sh.setColumnWidth(d * COLS_PER_DAY + 2, 95);
    sh.setColumnWidths(d * COLS_PER_DAY + 3, 3, 110);
  }
  sh.setFrozenRows(2);
  return 'rendered ' + totalWeeks + ' weeks';
}

// ====================== PART 3: EMAIL NOTIFY ======================
/** ส่งสรุปตารางวันพรุ่งนี้ (ข้ามถ้าพรุ่งนี้เป็นเสาร์/อาทิตย์และไม่มีรายการ) */
function sendTomorrowDigest() {
  const digest = tomorrowDigest_(new Date());
  if (digest) sendMail_(digest.subject, digest.html, digest.day);
}

function tomorrowDigest_(now) {
  const tomorrow = new Date(fmtDate_(now) + 'T00:00:00Z');
  tomorrow.setUTCDate(tomorrow.getUTCDate() + 1);
  const dt = tomorrow.toISOString().slice(0, 10);
  const [day] = buildDays_([dt]);
  const hasAny = day.staff.some(s => Object.values(s.shifts).some(a => a.length)) || day.trainee.some(x => x.items.length);
  if (!hasAny && (day.dow === 0 || day.dow === 6)) return null;

  const subject = `[${getConfig_('ORG_NAME', '')}] ตารางแพทย์วันพรุ่งนี้ ${day.label} ${tomorrow.getUTCFullYear() + 543}`;
  return { subject, html: buildDigestHtml_('ตารางการทำงานวันพรุ่งนี้', [day]), day };
}

/** ส่งสรุปสัปดาห์หน้า (จันทร์–ศุกร์) — รันเช้าวันอาทิตย์ */
function sendWeeklyDigest() {
  const t = new Date(); t.setDate(t.getDate() + 1); // อาทิตย์ +1 = จันทร์
  const week = apiGetWeek(fmtDate_(t), 5);
  const subject = `[${getConfig_('ORG_NAME', '')}] ตารางแพทย์สัปดาห์หน้า ${thaiDate_(week.start)} – ${thaiDate_(week.end)}`;
  const html = buildDigestHtml_('ตารางการทำงานสัปดาห์หน้า', week.days);
  sendMail_(subject, html, null);
}

function buildDigestHtml_(title, days) {
  const url = getConfig_('WEBAPP_URL', '');
  let h = `<div style="font-family:Arial,'Helvetica Neue',sans-serif;font-size:14px;color:#222">
  <h2 style="margin:0 0 8px">${title}</h2>`;
  days.forEach(day => {
    h += `<h3 style="background:#dbe9f6;padding:6px 8px;margin:14px 0 4px;border-radius:4px">${day.label}</h3>
    <table cellpadding="5" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#ccc;width:100%">
      <tr style="background:#f0f0f0"><th style="text-align:left">ชื่อ</th><th>เช้า</th><th>บ่าย</th><th>อื่นๆ</th></tr>`;
    day.staff.forEach(s => {
      const cell = sh => s.shifts[sh].concat(s.shifts['ทั้งวัน']).map(x => esc_(label_(x)) + (x.Note ? ` <span style="color:#888">(${esc_(x.Note)})</span>` : '')).join('<br>') || '<span style="color:#bbb">–</span>';
      h += `<tr><td><b>${esc_(s.name)}</b></td><td>${cell('เช้า')}</td><td>${cell('บ่าย')}</td><td>${cell('อื่นๆ')}</td></tr>`;
    });
    h += `</table>`;
    const off = day.staff.map(s => {
      const xs = Object.values(s.shifts).flat().filter(isOffsite_);
      return xs.length ? `${esc_(s.name)} (${xs.map(x => esc_(x.Activity || x.Shift)).join(', ')})` : '';
    }).filter(Boolean);
    h += `<p style="margin:6px 0 0"><b>🚗 นอกสถานที่:</b> ${off.length ? off.join(' · ') : '<span style="color:#888">ไม่มีรายการนอกสถานที่</span>'}</p>`;
    const tr = day.trainee.filter(x => x.items.length);
    if (tr.length) {
      h += `<p style="margin:6px 0 0"><b>แพทย์ elective:</b> ` +
        tr.map(x => `${x.type} ${x.items.map(i => esc_(i.Institution)).filter(Boolean).join('+')} (${x.items.reduce((a, i) => a + (Number(i.Count) || 0), 0)})`).join(' · ') + `</p>`;
    }
  });
  if (url) h += `<p style="margin-top:16px"><a href="${url}" style="background:#1a73e8;color:#fff;padding:8px 14px;border-radius:4px;text-decoration:none">เปิดตาราง / แก้ไข</a></p>`;
  h += `<p style="color:#999;font-size:12px">ส่งอัตโนมัติจากระบบตารางแพทย์ · ${Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm')}</p></div>`;
  return h;
}

function sendMail_(subject, html, dayOrNull) {
  const admins = String(getConfig_('ADMIN_EMAILS', '')).split(',').map(s => s.trim()).filter(Boolean);
  const to = new Set(admins);
  if (String(getConfig_('SEND_TO_STAFF', 'true')) === 'true') {
    getStaff_().forEach(s => { if (s.Email) to.add(String(s.Email).trim()); });
  }
  if (!to.size) { Logger.log('no recipients'); return; }
  MailApp.sendEmail({ to: [...to].join(','), subject, htmlBody: html, name: getConfig_('ORG_NAME', 'ตารางแพทย์') });
  Logger.log('sent to %s', [...to].join(','));
}

// ====================== HELPERS ======================
function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต ' + name + ' — กรุณารัน setup() ก่อน');
  return sh;
}

function getOrCreateSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (header && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#dbe9f6');
    sh.setFrozenRows(1);
  } else if (header) {
    // migration: เพิ่มคอลัมน์ใหม่ที่ยังไม่มี (เช่น Location) ต่อท้าย
    const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    header.forEach(h => {
      if (cur.indexOf(h) < 0) {
        const c = sh.getLastColumn() + 1;
        sh.getRange(1, c).setValue(h).setFontWeight('bold').setBackground('#dbe9f6');
        cur.push(h);
      }
    });
  }
  return sh;
}

function readTable_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values[0];
  return values.slice(1).filter(r => r.some(v => v !== '')).map((r, i) => {
    const o = { _row: i + 2 };
    header.forEach((h, j) => {
      let v = r[j];
      if (v instanceof Date) v = h === 'Date' ? fmtDate_(v) : v.toISOString();
      o[h] = v;
    });
    return o;
  });
}

function getStaff_() {
  return readTable_(SHEETS.STAFF)
    .filter(s => s.Name && String(s.Active).toLowerCase() !== 'false' && s.Active !== false)
    .sort((a, b) => (Number(a.Order) || 99) - (Number(b.Order) || 99));
}

function getSchedule_(from, to) {
  return readTable_(SHEETS.SCHEDULE).filter(x => x.Date >= from && x.Date <= to);
}

function getTrainee_(from, to) {
  return readTable_(SHEETS.TRAINEE).filter(x => x.Date >= from && x.Date <= to);
}

function getConfig_(key, def) {
  try {
    const row = readTable_(SHEETS.CONFIG).find(r => r.Key === key);
    return row && row.Value !== '' ? row.Value : def;
  } catch (e) { return def; }
}

function findRowById_(sh, id) {
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function validate_(o, keys) {
  keys.forEach(k => { if (!o || o[k] === undefined || o[k] === null || String(o[k]).trim() === '') throw new Error('กรุณากรอก ' + k); });
}

/** yyyy-MM-dd ใน timezone กรุงเทพ */
function fmtDate_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }

function parseDate_(s) {
  if (s instanceof Date) return s;
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) throw new Error('รูปแบบวันที่ไม่ถูกต้อง: ' + s);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** วันจันทร์ของสัปดาห์ */
function weekStart_(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const dow = (x.getDay() + 6) % 7; // Mon=0
  x.setDate(x.getDate() - dow);
  return x;
}

function thaiDate_(s) {
  const d = parseDate_(s);
  return d.getDate() + ' ' + THAI_MONTHS[d.getMonth()] + ' ' + (d.getFullYear() + 543);
}

function isOffsite_(x) { return x && x.Location === 'นอกสถานที่'; }

/** ข้อความในช่องตาราง: ใส่ 🚗 นำหน้าถ้านอกสถานที่ */
function label_(x) {
  const label = (isOffsite_(x) ? '🚗 ' : '') + (x.Activity || x.Shift);
  return x.MeetingType ? label + ' (' + x.MeetingType + ')' : label;
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
