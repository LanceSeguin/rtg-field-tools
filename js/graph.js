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

  // ── Calendars ─────────────────────────────────────────────────────────────

  /**
   * Returns all calendars for the signed-in user.
   * If shared calendar is enabled in settings, also fetches the shared mailbox calendars.
   */
  async function getCalendars() {
    const data = await get('/me/calendars?$select=id,name,isDefaultCalendar,canEdit&$top=50');
    let cals = (data.value || []).map(c => ({ ...c, _shared: false }));

    // Shared calendar (requires Calendars.Read.Shared permission)
    const sharedEnabled = Settings.get('sharedCal', false);
    const sharedEmail   = Settings.get('sharedEmail', '').trim();

    if (sharedEnabled && sharedEmail) {
      try {
        const sh = await get(
          `/users/${encodeURIComponent(sharedEmail)}/calendars?$select=id,name&$top=20`
        );
        (sh.value || []).forEach(c => {
          cals.push({ ...c, _shared: true, _sharedEmail: sharedEmail });
        });
      } catch (e) {
        console.warn('Shared calendar unavailable:', e.message);
        // Non-fatal — user will see their own calendars only
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
    // Extend end by 1 day to make the range inclusive
    const endExtended = new Date(end);
    endExtended.setDate(endExtended.getDate() + 1);

    const s = start.toISOString();
    const e = endExtended.toISOString();

    // Fields we need — body content is needed for parsing
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
