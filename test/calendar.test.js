/**
 * Calendar all-day write-path tests (no network).
 *
 * Exercises applyICalChanges, which shares the bare-date / VALUE=DATE helpers
 * with createEvent's iCalendar builder.
 */

const assert = require('assert');
const {
  applyICalChanges,
  looksAllDay,
  icalDateValue,
  addDayToICalDate,
} = require('../calendar/caldav-client');

let passed = 0;
function test(name, fn) {
  try {
    fn();
    passed++;
    console.log(`  ok - ${name}`);
  } catch (err) {
    console.error(`  FAIL - ${name}`);
    console.error(err);
    process.exitCode = 1;
  }
}

const TIMED = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'BEGIN:VEVENT',
  'UID:abc@icloud-mcp',
  'DTSTART;TZID=America/Mexico_City:20260829T100000',
  'DTEND;TZID=America/Mexico_City:20260829T110000',
  'SUMMARY:Standup',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

console.log('calendar all-day write path');

test('explicit isAllDay rewrites DTSTART/DTEND as VALUE=DATE, drops TZID', () => {
  const out = applyICalChanges(TIMED, { start: '2026-08-29', end: '2026-08-29', isAllDay: true });
  assert.ok(out.includes('DTSTART;VALUE=DATE:20260829'), out);
  assert.ok(out.includes('DTEND;VALUE=DATE:20260830'), out); // DTEND exclusive
  assert.ok(!/DTSTART;TZID/.test(out), out);
  assert.ok(!/DTEND;TZID/.test(out), out);
});

test('bare YYYY-MM-DD start/end is auto-detected as all-day', () => {
  const out = applyICalChanges(TIMED, { start: '2026-08-29', end: '2026-08-31' });
  assert.ok(out.includes('DTSTART;VALUE=DATE:20260829'), out);
  assert.ok(out.includes('DTEND;VALUE=DATE:20260901'), out); // 31st inclusive -> Sep 1 exclusive
});

test('all-day with start only derives a single-day DTEND', () => {
  const out = applyICalChanges(TIMED, { start: '2026-08-29', isAllDay: true });
  assert.ok(out.includes('DTSTART;VALUE=DATE:20260829'), out);
  assert.ok(out.includes('DTEND;VALUE=DATE:20260830'), out);
});

test('isAllDay:false with bare dates stays a timed event', () => {
  const out = applyICalChanges(TIMED, { start: '2026-08-29', end: '2026-08-30', isAllDay: false });
  assert.ok(/DTSTART:20260829T000000Z/.test(out), out);
  assert.ok(!/VALUE=DATE/.test(out), out);
});

test('unrelated change leaves DTSTART/DTEND untouched', () => {
  const out = applyICalChanges(TIMED, { summary: 'Renamed' });
  assert.ok(out.includes('DTSTART;TZID=America/Mexico_City:20260829T100000'), out);
  assert.ok(out.includes('SUMMARY:Renamed'), out);
});

test('month/year rollover on DTEND', () => {
  const out = applyICalChanges(TIMED, { start: '2026-12-31', end: '2026-12-31', isAllDay: true });
  assert.ok(out.includes('DTEND;VALUE=DATE:20270101'), out);
});

test('looksAllDay: explicit flag wins over bare-date detection', () => {
  assert.strictEqual(looksAllDay({ isAllDay: true, start: '2026-08-29T10:00:00' }), true);
  assert.strictEqual(looksAllDay({ isAllDay: false, start: '2026-08-29' }), false);
  assert.strictEqual(looksAllDay({ start: '2026-08-29' }), true);
  assert.strictEqual(looksAllDay({ start: '2026-08-29', end: '2026-08-30' }), true);
  assert.strictEqual(looksAllDay({ start: '2026-08-29T10:00:00' }), false);
  assert.strictEqual(looksAllDay({ start: '2026-08-29', end: '2026-08-30T10:00:00' }), false);
});

test('icalDateValue normalises bare / compact / ISO inputs without tz shift', () => {
  assert.strictEqual(icalDateValue('2026-08-29'), '20260829');
  assert.strictEqual(icalDateValue('20260829'), '20260829');
  assert.strictEqual(icalDateValue('2026-08-29T23:30:00-06:00'), '20260829');
});

test('addDayToICalDate rolls month and year', () => {
  assert.strictEqual(addDayToICalDate('20260829'), '20260830');
  assert.strictEqual(addDayToICalDate('20260831'), '20260901');
  assert.strictEqual(addDayToICalDate('20261231'), '20270101');
});

console.log(`\n${passed} passed`);
