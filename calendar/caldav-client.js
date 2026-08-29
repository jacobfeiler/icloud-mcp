/**
 * CalDAV client for iCloud Calendar
 */

const { DAVClient } = require('tsdav');
const ICAL = require('ical.js');
const config = require('../config');
const { getCredentials } = require('../auth');

let cachedClient = null;

/**
 * Get or create CalDAV client
 */
async function getClient() {
  if (cachedClient) {
    return cachedClient;
  }

  const creds = getCredentials();

  const client = new DAVClient({
    serverUrl: config.CALDAV.SERVER_URL,
    credentials: {
      username: creds.email,
      password: creds.password
    },
    authMethod: 'Basic',
    defaultAccountType: 'caldav'
  });

  try {
    await client.login();
    cachedClient = client;
    return client;
  } catch (error) {
    if (error.message?.includes('401') || error.message?.includes('auth')) {
      throw new Error('UNAUTHORIZED');
    }
    throw error;
  }
}

/**
 * Clear cached client (for re-auth)
 */
function clearClient() {
  cachedClient = null;
}

/**
 * Get all calendars
 */
async function getCalendars() {
  const client = await getClient();
  const calendars = await client.fetchCalendars();
  return calendars.map(cal => ({
    url: cal.url,
    displayName: cal.displayName || 'Unnamed Calendar',
    ctag: cal.ctag,
    syncToken: cal.syncToken
  }));
}

/**
 * Parse iCalendar event to simple object
 */
function parseEvent(icalData, url) {
  try {
    const jcalData = ICAL.parse(icalData);
    const comp = new ICAL.Component(jcalData);
    const vevent = comp.getFirstSubcomponent('vevent');

    if (!vevent) return null;

    const event = new ICAL.Event(vevent);

    return {
      url,
      uid: event.uid,
      summary: event.summary || '(No title)',
      description: event.description || '',
      location: event.location || '',
      start: event.startDate?.toJSDate(),
      end: event.endDate?.toJSDate(),
      isAllDay: event.startDate?.isDate || false,
      organizer: vevent.getFirstPropertyValue('organizer'),
      attendees: vevent.getAllProperties('attendee').map(a => a.getFirstValue()),
      status: event.status,
      created: vevent.getFirstPropertyValue('created')?.toJSDate(),
      lastModified: vevent.getFirstPropertyValue('last-modified')?.toJSDate()
    };
  } catch (error) {
    console.error('Error parsing event:', error.message);
    return null;
  }
}

/**
 * Expand one VCALENDAR object into the actual event occurrences that fall
 * inside [windowStart, windowEnd].
 *
 * A CalDAV time-range REPORT returns the *master* VEVENT for a recurring
 * series (DTSTART = the original, possibly years ago) plus any override
 * VEVENTs. Reading `startDate` off the master shows the wrong date and lets
 * long-past series leak in, so recurring masters are stepped with ical.js's
 * iterator and only in-window instances are emitted. Non-recurring events
 * are range-checked directly.
 *
 * @returns {Array<{uid,summary,description,location,start:Date,end:Date,isAllDay,status}>}
 */
function expandOccurrences(icalData, windowStart, windowEnd) {
  const MAX_ITER = 3000;
  try {
    const comp = new ICAL.Component(ICAL.parse(icalData));
    const vevents = comp.getAllSubcomponents('vevent');
    if (!vevents.length) return [];

    const master = vevents.find(v => !v.hasProperty('recurrence-id')) || vevents[0];
    const event = new ICAL.Event(master);
    for (const v of vevents) {
      if (v.hasProperty('recurrence-id')) {
        try { event.relateException(new ICAL.Event(v)); } catch (_) { /* ignore */ }
      }
    }

    const base = {
      uid: event.uid,
      summary: event.summary || '(No title)',
      description: event.description || '',
      location: event.location || '',
      isAllDay: event.startDate ? event.startDate.isDate : false,
      status: event.status
    };

    if (!event.isRecurring || !event.isRecurring()) {
      const start = event.startDate ? event.startDate.toJSDate() : null;
      const end = event.endDate ? event.endDate.toJSDate() : start;
      if (!start) return [];
      if (end < windowStart || start > windowEnd) return [];
      return [{ ...base, start, end }];
    }

    const out = [];
    const iter = event.iterator();
    let next;
    let i = 0;
    while ((next = iter.next()) && i++ < MAX_ITER) {
      const startJs = next.toJSDate();
      if (startJs > windowEnd) break;
      if (startJs < windowStart) continue;
      let d = null;
      try { d = event.getOccurrenceDetails(next); } catch (_) { /* use raw */ }
      out.push({
        ...base,
        summary: (d && d.item && d.item.summary) || base.summary,
        start: d ? d.startDate.toJSDate() : startJs,
        end: d ? d.endDate.toJSDate() : startJs
      });
      if (out.length >= 100) break;
    }
    return out;
  } catch (error) {
    console.error('expandOccurrences error:', error.message);
    return [];
  }
}

/**
 * List events from all calendars, recurring series expanded to the
 * occurrences inside the look-ahead window.
 */
async function listEvents(count = 25, daysAhead = 30) {
  const client = await getClient();
  const calendars = await client.fetchCalendars();

  const now = new Date();
  const endDate = new Date(now.getTime() + daysAhead * 24 * 60 * 60 * 1000);

  const allEvents = [];

  for (const calendar of calendars) {
    let calendarObjects;
    try {
      calendarObjects = await client.fetchCalendarObjects({
        calendar,
        timeRange: { start: now.toISOString(), end: endDate.toISOString() }
      });
    } catch (error) {
      console.error(`Error fetching from calendar ${calendar.displayName}:`, error.message);
      continue;
    }

    for (const obj of calendarObjects) {
      for (const occ of expandOccurrences(obj.data, now, endDate)) {
        occ.url = obj.url;
        occ.calendarName = calendar.displayName || 'Calendar';
        allEvents.push(occ);
      }
    }
  }

  allEvents.sort((a, b) => (a.start || 0) - (b.start || 0));

  return allEvents.slice(0, count);
}

/**
 * A bare calendar date with no time component, e.g. "2026-08-29".
 */
function isBareDate(value) {
  return typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value.trim());
}

/**
 * Decide whether an event should be written as a true all-day event.
 * Explicit `isAllDay` wins; otherwise a bare YYYY-MM-DD start (with a bare or
 * absent end) is treated as all-day.
 */
function looksAllDay({ isAllDay, start, end }) {
  if (isAllDay === true) return true;
  if (isAllDay === false) return false;
  return isBareDate(start) && (end === undefined || end === null || isBareDate(end));
}

/**
 * Reduce any date input ("2026-08-29", "20260829", or an ISO datetime) to the
 * iCalendar DATE value "YYYYMMDD", using the calendar day as written (no
 * timezone shift).
 */
function icalDateValue(input) {
  const s = String(input).trim();
  const m = s.match(/^(\d{4})-?(\d{2})-?(\d{2})/);
  if (m) return `${m[1]}${m[2]}${m[3]}`;
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    return `${d.getUTCFullYear()}${String(d.getUTCMonth() + 1).padStart(2, '0')}${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  throw new Error(`Cannot parse date: ${input}`);
}

/**
 * Add one day to a "YYYYMMDD" DATE value. Per RFC 5545 an all-day DTEND is
 * exclusive, so a single-day event ending on day N carries DTEND = N + 1.
 */
function addDayToICalDate(ymd) {
  const d = new Date(Date.UTC(+ymd.slice(0, 4), +ymd.slice(4, 6) - 1, +ymd.slice(6, 8) + 1));
  return `${d.getUTCFullYear()}${String(d.getUTCMonth() + 1).padStart(2, '0')}${String(d.getUTCDate()).padStart(2, '0')}`;
}

/**
 * Create a new event
 */
async function createEvent({ summary, start, end, description, location, calendarUrl, isAllDay }) {
  const client = await getClient();

  // Get calendars if URL not provided
  let targetCalendar;
  if (calendarUrl) {
    const calendars = await client.fetchCalendars();
    targetCalendar = calendars.find(c => c.url === calendarUrl);
  }

  if (!targetCalendar) {
    const calendars = await client.fetchCalendars();
    targetCalendar = calendars[0]; // Use first calendar
  }

  if (!targetCalendar) {
    throw new Error('No calendar found');
  }

  // Create iCalendar data
  const uid = `${Date.now()}-${Math.random().toString(36).substr(2, 9)}@icloud-mcp`;

  let dtStartLine;
  let dtEndLine;
  if (looksAllDay({ isAllDay, start, end })) {
    // True all-day: DATE-valued DTSTART/DTEND, DTEND exclusive (day after the
    // last day). `end` is the inclusive last day and defaults to a single day.
    const startYmd = icalDateValue(start);
    const endYmd = (end === undefined || end === null)
      ? addDayToICalDate(startYmd)
      : addDayToICalDate(icalDateValue(end));
    dtStartLine = `DTSTART;VALUE=DATE:${startYmd}`;
    dtEndLine = `DTEND;VALUE=DATE:${endYmd}`;
  } else {
    dtStartLine = `DTSTART:${formatICalDate(new Date(start))}`;
    dtEndLine = `DTEND:${formatICalDate(new Date(end))}`;
  }

  const icalData = `BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//iCloud MCP//EN
BEGIN:VEVENT
UID:${uid}
DTSTAMP:${formatICalDate(new Date())}
${dtStartLine}
${dtEndLine}
SUMMARY:${escapeICalText(summary)}${description ? `\nDESCRIPTION:${escapeICalText(description)}` : ''}${location ? `\nLOCATION:${escapeICalText(location)}` : ''}
END:VEVENT
END:VCALENDAR`;

  const result = await client.createCalendarObject({
    calendar: targetCalendar,
    filename: `${uid}.ics`,
    iCalString: icalData
  });

  return {
    success: true,
    uid,
    url: result?.url,
    calendar: targetCalendar.displayName
  };
}

/**
 * Rewrite individual properties of an existing VEVENT.
 *
 * Deliberately a line-level edit rather than rebuilding the iCalendar from
 * scratch: a rebuild would silently drop RRULE, ATTENDEE, VALARM and any other
 * property this server does not model, which on a real calendar means losing
 * recurrence and invitees on every edit.
 *
 * @param {string} ical - Existing iCalendar text
 * @param {Object} changes - Properties to set; undefined values are left alone
 * @returns {string} - Updated iCalendar text
 */
function applyICalChanges(ical, { summary, start, end, description, location, isAllDay }) {
  // All-day when asked explicitly, or (absent an explicit flag) when a bare
  // YYYY-MM-DD start/end is supplied. Switching an existing timed event to
  // all-day this way requires passing start and end together.
  const allDay = isAllDay === true ||
    (isAllDay === undefined && (isBareDate(start) || isBareDate(end)));

  const updates = [];
  if (summary !== undefined) updates.push(['SUMMARY', escapeICalText(summary)]);
  if (allDay) {
    if (start !== undefined) updates.push(['DTSTART', icalDateValue(start), ';VALUE=DATE']);
    if (start !== undefined || end !== undefined) {
      const endSource = (end === undefined || end === null) ? start : end;
      updates.push(['DTEND', addDayToICalDate(icalDateValue(endSource)), ';VALUE=DATE']);
    }
  } else {
    if (start !== undefined) updates.push(['DTSTART', formatICalDate(new Date(start))]);
    if (end !== undefined) updates.push(['DTEND', formatICalDate(new Date(end))]);
  }
  if (description !== undefined) updates.push(['DESCRIPTION', escapeICalText(description)]);
  if (location !== undefined) updates.push(['LOCATION', escapeICalText(location)]);
  updates.push(['DTSTAMP', formatICalDate(new Date())]);

  const lines = ical.split(/\r?\n/);

  for (const [key, value, params = ''] of updates) {
    // A property may carry parameters, e.g. DTSTART;TZID=Europe/Madrid:... —
    // the whole line is replaced, so a timed->all-day switch drops the TZID.
    const idx = lines.findIndex(l => {
      const upper = l.toUpperCase();
      return upper.startsWith(key + ':') || upper.startsWith(key + ';');
    });

    if (idx !== -1) {
      lines[idx] = `${key}${params}:${value}`;
    } else {
      const endIdx = lines.findIndex(l => l.toUpperCase().startsWith('END:VEVENT'));
      lines.splice(endIdx === -1 ? lines.length : endIdx, 0, `${key}${params}:${value}`);
    }
  }

  return lines.join('\r\n');
}

/**
 * Update an existing event, preserving every property not being changed.
 */
async function updateEvent(eventUrl, changes) {
  const client = await getClient();
  const calendars = await client.fetchCalendars();

  let existing = null;
  for (const calendar of calendars) {
    try {
      const objects = await client.fetchCalendarObjects({
        calendar,
        objectUrls: [eventUrl]
      });
      if (objects && objects.length && objects[0].data) {
        existing = objects[0];
        break;
      }
    } catch (error) {
      // Wrong calendar for this URL; keep looking.
    }
  }

  if (!existing) {
    throw new Error(`Event not found: ${eventUrl}`);
  }

  const updated = applyICalChanges(existing.data, changes);

  await client.updateCalendarObject({
    calendarObject: {
      url: eventUrl,
      data: updated,
      etag: existing.etag || ''
    }
  });

  return { success: true, url: eventUrl };
}

/**
 * Delete an event
 */
async function deleteEvent(eventUrl) {
  const client = await getClient();

  await client.deleteCalendarObject({
    calendarObject: {
      url: eventUrl,
      etag: '' // Will be fetched
    }
  });

  return { success: true };
}

/**
 * Format date for iCalendar
 */
function formatICalDate(date) {
  return date.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
}

/**
 * Escape text for iCalendar
 */
function escapeICalText(text) {
  if (!text) return '';
  return text
    .replace(/\\/g, '\\\\')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,')
    .replace(/\n/g, '\\n');
}

module.exports = {
  getClient,
  clearClient,
  getCalendars,
  listEvents,
  createEvent,
  updateEvent,
  deleteEvent,
  parseEvent,
  expandOccurrences,
  applyICalChanges,
  looksAllDay,
  icalDateValue,
  addDayToICalDate
};
