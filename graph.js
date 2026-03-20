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

  // ── Calendar name filter ───────────────────────────────────────────────────
  // Calendars whose names contain any of these strings are hidden.
  // Add to this list if other junk calendars appear.
  const SKIP_NAMES = [
    'birthday', 'birthdays',
    'holiday', 'holidays',
    'united states holidays',
    'other people',
  ];

  function _shouldSkip(name) {
    const lower = (name || '').toLowerCase().trim();
    return SKIP_NAMES.some(s => lower.includes(s));
  }

  // ── Calendars ─────────────────────────────────────────────────────────────

  /**
   * Returns only the user's DEFAULT calendar, plus the shared calendar
   * if enabled in settings.
   *
   * We deliberately skip secondary personal calendars (Birthdays, Holidays,
   * etc.) — users should only see their main calendar and the shared one.
   */
  async function getCalendars() {
    const data = await get('/me/calendars?$select=id,name,isDefaultCalendar,canEdit&$top=50');
    const all  = data.value || [];

    // Keep ONLY the default calendar from the user's own mailbox
    let cals = all
      .filter(c => c.isDefaultCalendar === true && !_shouldSkip(c.name))
      .map(c => ({ ...c, _shared: false }));

    // Fallback: if somehow no default was flagged, take the first non-junk one
    if (!cals.length) {
      const first = all.find(c => !_shouldSkip(c.name));
      if (first) cals = [{ ...first, _shared: false }];
    }

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
