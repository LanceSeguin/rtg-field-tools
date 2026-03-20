// ─────────────────────────────────────────────────────────────────────────────
// graph.js — Microsoft Graph API calls
// Edit this file to change what calendar data is fetched or how it's shaped.
// ─────────────────────────────────────────────────────────────────────────────

const Graph = (() => {

  /** Generic GET against Graph API */
  async function get(url) {
    const r = await fetch('https://graph.microsoft.com/v1.0' + url, {
      headers: { Authorization: `Bearer ${Auth.getToken()}` }
    });
    if (!r.ok) {
      const body = await r.text();
      throw new Error(`Graph ${r.status}: ${body}`);
    }
    return r.json();
  }

  // ── Calendar filter ────────────────────────────────────────────────────────
  // Skip calendars whose names contain any of these strings (case-insensitive).
  // Add more entries here if other unwanted calendars keep appearing.
  const SKIP_NAMES = [
    'birthday', 'birthdays',
    'holiday', 'holidays',
    'united states',
    'other people',
  ];

  function _shouldSkip(name) {
    const lower = (name || '').toLowerCase().trim();
    // Skip by keyword
    if (SKIP_NAMES.some(s => lower.includes(s))) return true;
    // Skip calendars with [XXX/YYY] bracket notation — these are other people's
    // calendars shared to you (e.g. "Longoria, Alfredo [EMR/MSOL/HSTN]")
    if (/\[.+\/.+\]/.test(name || '')) return true;
    return false;
  }

  // ── Calendars ─────────────────────────────────────────────────────────────

  /**
   * Returns only the user's primary calendar, plus the shared service calendar
   * if enabled in Settings. Everything else is filtered out.
   */
  async function getCalendars() {
    const data = await get('/me/calendars?$select=id,name,isDefaultCalendar,canEdit&$top=50');
    const all  = data.value || [];

    // Strategy 1: calendar explicitly flagged as default by Graph
    let primary = all.find(c => c.isDefaultCalendar === true);

    // Strategy 2: calendar literally named "Calendar" (common when flag isn't set)
    if (!primary) primary = all.find(c => c.name === 'Calendar');

    // Strategy 3: first calendar that isn't junk
    if (!primary) primary = all.find(c => !_shouldSkip(c.name));

    let cals = primary ? [{ ...primary, _shared: false }] : [];

    // Shared calendar (only if enabled in Settings AND address is set)
    const sharedEnabled = Settings.get('sharedCal', false);
    const sharedEmail   = Settings.get('sharedEmail', '').trim();

    if (sharedEnabled && sharedEmail) {
      try {
        const sh = await get(
          `/users/${encodeURIComponent(sharedEmail)}/calendars?$select=id,name,isDefaultCalendar&$top=20`
        );
        const sharedDefault = (sh.value || []).find(c => c.isDefaultCalendar);
        const sharedCal     = sharedDefault || (sh.value || [])[0];
        if (sharedCal) {
          cals.push({
            ...sharedCal,
            name:         sharedEmail,
            _shared:      true,
            _sharedEmail: sharedEmail,
          });
        }
      } catch (e) {
        console.warn('Shared calendar unavailable:', e.message);
      }
    }

    return cals;
  }

  // ── Events ────────────────────────────────────────────────────────────────

  /**
   * Fetch events from a calendar within a date range.
   * @param {string} calendarId   - Graph calendar ID
   * @param {Date}   start        - range start
   * @param {Date}   end          - range end (inclusive)
   * @param {string} sharedEmail  - if set, fetches from a shared user's calendar
   */
  async function getEvents(calendarId, start, end, sharedEmail) {
    const s = start.toISOString();
    const endExtended = new Date(end);
    endExtended.setDate(endExtended.getDate() + 1);
    const e = endExtended.toISOString();

    const select  = 'subject,body,bodyPreview,start,end,location,organizer';
    const filter  = `start/dateTime ge '${s}' and end/dateTime le '${e}'`;
    const orderby = 'start/dateTime';
    const top     = 100;

    const base = sharedEmail
      ? `/users/${encodeURIComponent(sharedEmail)}/calendars/${calendarId}/events`
      : `/me/calendars/${calendarId}/events`;

    const url = `${base}?$select=${select}&$filter=${encodeURIComponent(filter)}&$orderby=${orderby}&$top=${top}`;
    const data = await get(url);
    return data.value || [];
  }

  return { get, getCalendars, getEvents };
})();
