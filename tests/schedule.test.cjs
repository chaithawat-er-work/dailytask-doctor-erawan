const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { test } = require('node:test');
const { runInNewContext } = require('node:vm');

const source = readFileSync(join(__dirname, '../src/Code.gs'), 'utf8');
const scheduleHeader = ['ID', 'Date', 'Name', 'Shift', 'Activity', 'Note', 'RecordedBy', 'Timestamp', 'Location'];
const meetingTypes = [
  ['ประชุมในเอราวัณประจำ', 'ในศูนย์เอราวัณ'],
  ['ประชุมในเอราวัณไม่ประจำ', 'ในศูนย์เอราวัณ'],
  ['ประชุมนอกเอราวัณประจำ', 'นอกสถานที่'],
  ['ประชุมนอกเอราวัณไม่ประจำ', 'นอกสถานที่'],
  ['ประชุมออนไลน์ประจำ', 'ออนไลน์'],
];

function environment() {
  const tables = {
    Schedule: [scheduleHeader.slice()],
    Staff: [['Name', 'Role', 'Email', 'Active', 'Order'], ['แพทย์ 1', 'แพทย์', '', true, 1], ['แพทย์ 2', 'แพทย์', '', true, 2]],
    Trainee: [['ID', 'Date', 'Type', 'Institution', 'Count', 'Note', 'RecordedBy', 'Timestamp']],
    Config: [['Key', 'Value', 'Description'], ['ORG_NAME', 'ชุดทดสอบ', '']],
  };
  function sheet(name) {
    const rows = tables[name];
    return {
      getDataRange: () => ({ getValues: () => rows }),
      getLastColumn: () => Math.max(...rows.map(row => row.length)),
      getMaxColumns: () => 26,
      getLastRow: () => rows.length,
      getMaxRows: () => 100,
      appendRow: row => rows.push(row),
      deleteRow: row => rows.splice(row - 1, 1),
      getRange(row, col, count = 1, width = 1) {
        if (row === 'A2:A') return {};
        const range = {
          getValues: () => Array.from({ length: count }, (_, i) =>
            Array.from({ length: width }, (_, j) => rows[row + i - 1]?.[col + j - 1] ?? '')),
          getValue: () => range.getValues()[0][0],
          setValues(values) {
            values.forEach((cells, i) => {
              const target = rows[row + i - 1] ||= [];
              cells.forEach((value, j) => { target[col + j - 1] = value; });
            });
            return range;
          },
          setValue: value => range.setValues([[value]]),
          setFontWeight: () => range,
          setBackground: () => range,
          setDataValidation() {},
        };
        return range;
      },
    };
  }
  const validation = {
    requireValueInRange: () => validation,
    requireValueInList: () => validation,
    setAllowInvalid: () => validation,
    build: () => validation,
  };
  const user = () => ({ getEmail: () => 'tester@example.invalid' });
  let nextId = 0;
  const api = runInNewContext(source + '\n({ setupMeetingTypes, apiGetMeta, apiSaveSchedule, apiSaveScheduleRange, apiSaveScheduleBulk, apiDeleteSchedule, apiGetWeek, apiGetMonth, tomorrowDigest_ });', {
    Date,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: sheet }), newDataValidation: () => validation },
    Session: { getActiveUser: user, getEffectiveUser: user },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      getUuid: () => String(++nextId),
      formatDate(date, zone, format) {
        const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
          timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23',
        }).formatToParts(date).map(part => [part.type, part.value]));
        return format === 'yyyy-MM-dd'
          ? [parts.year, parts.month, parts.day].join('-')
          : [parts.day, parts.month, parts.year].join('/') + ' ' + parts.hour + ':' + parts.minute;
      },
    },
    MailApp: { sendEmail() { assert.fail('Tests must never send mail'); } },
  });
  return { api, tables };
}

const baseItem = { date: '2026-11-18', name: 'แพทย์ 1', shift: 'เช้า', activity: 'ทบทวนงาน' };

test('daily digest selects Wednesday, not Monday, across Bangkok midnight', () => {
  const { api } = environment();
  api.apiSaveSchedule({ ...baseItem, date: '2026-11-16', activity: 'รายการวันจันทร์' });
  api.apiSaveSchedule({ ...baseItem, activity: 'รายการวันพุธ' });
  const digest = api.tomorrowDigest_(new Date('2026-11-16T17:15:00Z'));
  assert.equal(digest.day.date, '2026-11-18');
  assert.match(digest.subject, /พุธ 18 พ\.ย\. 2569/);
  assert.match(digest.html, /รายการวันพุธ/);
  assert.doesNotMatch(digest.html, /รายการวันจันทร์/);
});

test('daily digest handles weekends and New Year', () => {
  const { api, tables } = environment();
  api.apiSaveSchedule({ ...baseItem, date: '2026-11-16' });
  assert.equal(api.tomorrowDigest_(new Date('2026-11-20T10:00:00Z')), null);
  assert.equal(api.tomorrowDigest_(new Date('2026-11-21T10:00:00Z')), null);
  api.apiSaveSchedule({ ...baseItem, date: '2026-11-21' });
  assert.equal(api.tomorrowDigest_(new Date('2026-11-20T10:00:00Z')).day.date, '2026-11-21');
  tables.Trainee.push(['t1', new Date('2026-11-21T17:00:00Z'), 'Resident', 'สถาบันทดสอบ', 1]);
  const sunday = api.tomorrowDigest_(new Date('2026-11-21T10:00:00Z'));
  assert.equal(sunday.day.date, '2026-11-22');
  assert.match(sunday.html, /สถาบันทดสอบ/);
  assert.match(api.tomorrowDigest_(new Date('2026-12-31T10:00:00Z')).subject, /ศุกร์ 1 ม\.ค\. 2570/);
});

test('migration preserves old rows, is repeatable and rejects conflicting columns', () => {
  const { api, tables } = environment();
  const previous = api.apiSaveSchedule(baseItem);
  const oldRow = Array.from(tables.Schedule[1]);
  assert.equal(api.apiGetMeta().meetingTypesEnabled, false);
  assert.throws(() => api.apiSaveSchedule({ ...baseItem, meetingType: meetingTypes[0][0] }), /เปิดใช้ประเภท/);
  tables.Schedule[0].push('ExistingExtra');
  assert.throws(() => api.setupMeetingTypes(), /คอลัมน์เพิ่มเติม/);
  assert.equal(tables.Schedule[0][9], 'ExistingExtra');
  tables.Schedule[0].pop();
  api.setupMeetingTypes();
  api.setupMeetingTypes();
  assert.deepEqual(tables.Schedule[0], scheduleHeader.concat('MeetingType'));
  assert.deepEqual(Array.from(tables.Schedule[1]), oldRow);
  assert.equal(api.apiGetMeta().meetingTypesEnabled, true);
  api.apiSaveSchedule({ ...baseItem, id: previous.id, activity: 'แก้ไขงานเดิม' });
  assert.equal(tables.Schedule[1][9], '');
  assert.throws(() => api.apiSaveSchedule({ ...baseItem, meetingType: '__proto__' }), /ประเภทการประชุมไม่ถูกต้อง/);
});

test('saves all five meeting types and derives their locations', () => {
  const { api, tables } = environment();
  api.setupMeetingTypes();
  for (const [meetingType] of meetingTypes) api.apiSaveSchedule({ ...baseItem, meetingType });
  assert.deepEqual(tables.Schedule.slice(1).map(row => [row[9], row[8]]), meetingTypes);
  const digest = api.tomorrowDigest_(new Date('2026-11-17T10:00:00Z'));
  assert.match(digest.html, /ประชุมออนไลน์ประจำ/);
  const week = api.apiGetWeek(baseItem.date, 5);
  const month = api.apiGetMonth(baseItem.date);
  for (const view of [week, month]) {
    const day = view.days.find(day => day.date === baseItem.date);
    assert.equal(day.staff[0].shifts['เช้า'][4].MeetingType, 'ประชุมออนไลน์ประจำ');
  }
});

test('bulk and range save both dates, edit preserves or clears the category, delete keeps other rows', () => {
  const { api, tables } = environment();
  api.setupMeetingTypes();
  const dates = { dateFrom: '2026-11-18', dateTo: '2026-11-19' };
  const meetingType = 'ประชุมออนไลน์ประจำ';
  const bulk = api.apiSaveScheduleBulk({ ...baseItem, ...dates, names: ['แพทย์ 1', 'แพทย์ 2'], meetingType });
  assert.equal(bulk.count, 4);
  const range = api.apiSaveScheduleRange({ ...baseItem, ...dates, meetingType });
  assert.equal(range.count, 2);
  const id = tables.Schedule[1][0];
  api.apiSaveSchedule({ ...baseItem, id });
  assert.equal(tables.Schedule[1][9], meetingType);
  api.apiSaveSchedule({ ...baseItem, id, meetingType: '' });
  assert.equal(tables.Schedule[1][9], '');
  assert.equal(tables.Schedule[1][8], 'ในศูนย์เอราวัณ');
  assert.equal(api.apiDeleteSchedule(id).ok, true);
  assert.equal(tables.Schedule.length, 6);
  assert.ok(tables.Schedule.slice(1).every(row => row[9] === meetingType));
});

test('meeting selector, weekly and monthly views and export show the meeting category', async () => {
  const { api } = environment();
  api.setupMeetingTypes();
  api.apiSaveSchedule({ ...baseItem, meetingType: 'ประชุมออนไลน์ประจำ' });
  const mock = {
    apiGetMeta: { ...api.apiGetMeta(), today: baseItem.date },
    apiGetWeek: api.apiGetWeek(baseItem.date, 5),
    apiGetMonth: api.apiGetMonth(baseItem.date),
  };
  const html = readFileSync(join(__dirname, '../src/Index.html'), 'utf8');
  const [, script] = html.match(/<script>([\s\S]*?)<\/script>/);
  const elements = new Map();
  const document = {
    addEventListener() {},
    getElementById(id) {
      assert.ok(html.includes('id="' + id + '"'), 'Missing element: ' + id);
      if (!elements.has(id)) elements.set(id, { classList: { toggle() {} } });
      return elements.get(id);
    },
  };
  const app = runInNewContext(script.replace(/\binit\(\);\s*$/, '({ init, setView, loadWeek, buildExportHtml, syncMeetingLocation });'), {
    document,
    localStorage: { getItem() {}, setItem() {} },
    setTimeout() {},
    google: { script: { run: {
      withSuccessHandler: resolve => ({
        withFailureHandler: () => Object.fromEntries(Object.entries(mock).map(([key, value]) => [key, () => resolve(value)])),
      }),
    } } },
  });
  await app.init();
  await app.loadWeek(baseItem.date);
  for (const [type] of meetingTypes) assert.ok(elements.get('s_meeting').innerHTML.includes(type));
  elements.get('s_meeting').value = 'ประชุมออนไลน์ประจำ';
  app.syncMeetingLocation();
  assert.equal(elements.get('s_meetingLocation').textContent, 'ออนไลน์');
  assert.match(elements.get('grid').innerHTML, /ทบทวนงาน \(ประชุมออนไลน์ประจำ\)/);
  app.setView('month');
  await app.loadWeek(baseItem.date);
  assert.match(elements.get('mtable').innerHTML, /hb online/);
  assert.match(app.buildExportHtml(mock.apiGetWeek, '', 'week'), /ทบทวนงาน \(ประชุมออนไลน์ประจำ\)/);
});
