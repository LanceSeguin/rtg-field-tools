// ─────────────────────────────────────────────────────────────────────────────
// app.js — Main application controller
// Edit this file to add new screens, change navigation, or add new tools.
// ─────────────────────────────────────────────────────────────────────────────

// ── Persistent settings (localStorage) ───────────────────────────────────────
const Settings = {
  get: (key, def = '') => {
    try { const v = localStorage.getItem('rtg_' + key); return v !== null ? JSON.parse(v) : def; }
    catch { return def; }
  },
  set: (key, val) => {
    try { localStorage.setItem('rtg_' + key, JSON.stringify(val)); } catch {}
  },
};

// ── Calendar panel state ──────────────────────────────────────────────────────
const Cal = (() => {
  let _cals     = [];
  let _events   = [];
  let _selected = null;

  const DAYS = ['SUN','MON','TUE','WED','THU','FRI','SAT'];

  function _fmtDate(d) {
    return `${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}-${d.getFullYear()}`;
  }
  function _parseDate(s) {
    const [m,d,y] = s.split('-');
    return new Date(+y, +m-1, +d);
  }

  // Load calendars from Graph and populate the dropdown
  async function init() {
    try {
      _cals = await Graph.getCalendars();
      const sel = document.getElementById('cal-select');
      sel.innerHTML = _cals.map((c, i) =>
        `<option value="${i}">${c._shared ? '📤 ' : '📅 '}${c.name}</option>`
      ).join('');
    } catch (e) {
      App.toast('Could not load calendars: ' + e.message, 'err');
    }
  }

  // Set date range preset and reload events
  function preset(p) {
    const t = new Date();
    let s = new Date(t), e = new Date(t);
    if (p === 'week') {
      s.setDate(t.getDate() - ((t.getDay() + 6) % 7));
      e = new Date(s); e.setDate(s.getDate() + 6);
    } else if (p === 'month') {
      s = new Date(t.getFullYear(), t.getMonth(), 1);
      e = new Date(t.getFullYear(), t.getMonth() + 1, 0);
    }
    document.getElementById('cal-start').value = _fmtDate(s);
    document.getElementById('cal-end').value   = _fmtDate(e);
    load();
  }

  // Fetch events and render the list
  async function load() {
    const idx = parseInt(document.getElementById('cal-select').value) || 0;
    const cal = _cals[idx];
    if (!cal) return;

    const startS = document.getElementById('cal-start').value;
    const endS   = document.getElementById('cal-end').value;
    if (!startS || !endS) return;

    const list = document.getElementById('events-list');
    list.innerHTML = '<div class="events-loading">Loading events…</div>';
    _selected = null;

    try {
      const start = _parseDate(startS);
      const end   = _parseDate(endS);
      _events = await Graph.getEvents(cal.id, start, end, cal._sharedEmail);

      if (!_events.length) {
        list.innerHTML = '<div class="events-empty">No events found in this date range.</div>';
        return;
      }

      list.innerHTML = _events.map((ev, i) => {
        const s   = new Date(ev.start.dateTime || ev.start.date);
        const e   = new Date(ev.end.dateTime   || ev.end.date);
        const fmt = d => d.toLocaleDateString('en-US', {month:'short',day:'numeric',year:'numeric'});
        const ftm = d => d.toLocaleTimeString('en-US', {hour:'2-digit',minute:'2-digit'});
        const loc = ev.location?.displayName || '';
        return `<div class="event-card" id="ec-${i}" onclick="Cal.select(${i})">
          <div class="event-subj">${ev.subject || '(No subject)'}</div>
          <div class="event-when">${fmt(s)} ${ftm(s)} → ${fmt(e)} ${ftm(e)}</div>
          ${loc ? `<div class="event-loc">📍 ${loc}</div>` : ''}
        </div>`;
      }).join('');
    } catch (e) {
      list.innerHTML = `<div class="events-empty">Error loading events.<br><small>${e.message}</small></div>`;
      App.toast(e.message, 'err');
    }
  }

  function select(i) {
    document.querySelectorAll('.event-card').forEach(el => el.classList.remove('selected'));
    document.getElementById('ec-' + i)?.classList.add('selected');
    _selected = _events[i];
  }

  // Use the selected event — parse fields and populate form
  async function useSelected() {
    if (!_selected) { App.toast('Select an event first', 'warn'); return; }

    const ev   = _selected;
    const body = ev.body?.content || ev.bodyPreview || '';

    const fields = Parsing.parseEventFields(ev.subject || '', body);

    const set = (id, v) => { if (v) { const el = document.getElementById(id); if (el) el.value = v; } };
    set('f-customer', fields.CustomerName);
    set('f-agency',   fields.CustomerName);  // mirrors original app behavior
    set('f-tech',     fields.ServiceTechnician || Settings.get('defaultTech', CONFIG.defaultTechnician));
    set('f-order',    fields.ServiceAgencyOrder);
    set('f-cname',    fields.ContactName);
    set('f-cphone',   fields.ContactPhone);
    set('f-cemail',   fields.ContactEmail);
    set('f-po',       fields.PONumber);
    if (fields.ServiceAddress) document.getElementById('f-address').value = fields.ServiceAddress;
    if (fields.Scope)          document.getElementById('f-scope').value   = fields.Scope;

    // Auto-populate labor rows from event date range
    const start = new Date(ev.start.dateTime || ev.start.date);
    const end   = new Date(ev.end.dateTime   || ev.end.date);
    const dates = [];
    const cur   = new Date(start); cur.setHours(0, 0, 0, 0);
    const endDay= new Date(end);   endDay.setHours(0, 0, 0, 0);
    while (cur < endDay) { dates.push(new Date(cur)); cur.setDate(cur.getDate() + 1); }
    if (!dates.length) dates.push(start);
    Labor.populateFromDates(dates);

    App.toast('Form populated from calendar event ✓');
    App.tab('request');
  }

  return { init, preset, load, select, useSelected };
})();

// ── Main App controller ───────────────────────────────────────────────────────
const App = {

  // ── Boot ─────────────────────────────────────────────────────────────────
  async init() {
    // Handle OAuth redirect back from Microsoft
    if (location.hash.includes('code=')) {
      const ok = await Auth.handleRedirect();
      if (ok) { await App._enterApp(); return; }
    }
    // Restore existing session (page refresh within same tab)
    if (Auth.restoreSession()) { await App._enterApp(); return; }
    // No session — show login screen
    App._show('login');
  },

  // ── Auth ──────────────────────────────────────────────────────────────────
  async login()  { await Auth.login(); },
  async logout() { Auth.logout(); App._show('login'); },

  async _enterApp() {
    const user = await Auth.getUser();
    document.getElementById('hdr-username').textContent =
      user?.displayName || user?.mail || 'Signed in';

    // Restore saved settings into modal controls
    document.getElementById('tog-shared').checked        = Settings.get('sharedCal', false);
    document.getElementById('set-tech').value            = Settings.get('defaultTech', CONFIG.defaultTechnician);
    document.getElementById('set-shared-email').value   = Settings.get('sharedEmail', '');

    // Pre-fill tech field
    document.getElementById('f-tech').value = Settings.get('defaultTech', CONFIG.defaultTechnician);

    App._show('hub');
  },

  // ── Screen navigation ─────────────────────────────────────────────────────
  _show(screen) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    document.getElementById('screen-' + screen)?.classList.add('active');
  },

  showHub() { App._show('hub'); },

  async showWO() {
    App._show('wo');

    // Initialize rich text editors (idempotent — safe to call multiple times)
    RTE.init('rte-summary',  'rte-summary-tb');
    RTE.init('rte-followup', 'rte-followup-tb');

    // Set default calendar date range
    const t = new Date();
    const s = new Date(t); s.setDate(t.getDate() - CONFIG.calendarLookbackDays);
    const e = new Date(t); e.setDate(t.getDate() + CONFIG.calendarLookaheadDays);
    const fmt = d =>
      `${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}-${d.getFullYear()}`;
    document.getElementById('cal-start').value = fmt(s);
    document.getElementById('cal-end').value   = fmt(e);

    // Load calendars, fetch events, and load parts catalog in parallel
    await Promise.all([
      Cal.init().then(() => Cal.load()),
      Parts.loadCatalog(),
    ]);
  },

  // ── Tab navigation (within Work Order screen) ─────────────────────────────
  tab(name) {
    document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('tab-' + name)?.classList.add('active');
    document.querySelector(`.tab-btn[data-tab="${name}"]`)?.classList.add('active');
    if (name === 'output') App._updateSummary();
  },

  _updateSummary() {
    const g = id => document.getElementById(id)?.value?.trim() || '';
    const rows = [
      ['Customer',     g('f-customer')],
      ['PO Number',    g('f-po')],
      ['Agency Order', g('f-order')],
      ['Technician',   g('f-tech')],
      ['Product Line', g('f-product')],
      ['Serial',       g('f-serial')],
      ['Contact',      [g('f-cname'), g('f-cphone')].filter(Boolean).join(' · ')],
      ['Address',      g('f-address').replace(/\n/g, ', ')],
      ['Labor Rows',   String(Labor.count())],
      ['Parts Rows',   String(Parts.count())],
    ];
    document.getElementById('form-summary').innerHTML = rows
      .filter(([,v]) => v)
      .map(([k,v]) => `<div class="sum-key">${k}</div><div class="sum-val">${v}</div>`)
      .join('');
  },

  // ── Work order generation ─────────────────────────────────────────────────
  async generate() {
    const order = document.getElementById('f-order').value.trim();
    if (!order) { App.toast('Service Agency Order number is required', 'warn'); return; }

    const btn  = document.getElementById('gen-btn');
    const spin = document.getElementById('gen-spin');
    const lbl  = document.getElementById('gen-lbl');
    btn.disabled = true;
    spin.style.display = 'inline-block';
    lbl.textContent = 'Generating…';

    try {
      const g = id => document.getElementById(id)?.value || '';
      const data = {
        customerName:   g('f-customer'),
        poNumber:       g('f-po'),
        techName:       g('f-tech'),
        serviceOrder:   g('f-order'),
        productLine:    g('f-product'),
        systemSerial:   g('f-serial'),
        serviceAddress: g('f-address'),
        contactName:    g('f-cname'),
        contactPhone:   g('f-cphone'),
        contactEmail:   g('f-cemail'),
        scope:          g('f-scope'),
        laborRows:      Labor.getRows(),
        partsRows:      Parts.getRows(),
      };

      // Build filename from pattern
      const cust    = data.customerName.replace(/[^a-z0-9 \-_]/gi, '').trim().replace(/ /g, '_');
      const today   = new Date().toISOString().slice(0, 10);
      const pattern = document.getElementById('f-pattern').value || '{ServiceAgencyOrder}_Work Order.docx';
      const fname   = pattern
        .replace('{ServiceAgencyOrder}', order)
        .replace('{Customer}', cust)
        .replace('{Date}', today);

      await DOCX.download(data, fname.endsWith('.docx') ? fname : fname + '.docx');
      App.toast('✔ ' + fname + ' downloaded');
    } catch (e) {
      App.toast('Error: ' + e.message, 'err');
      console.error(e);
    } finally {
      btn.disabled = false;
      spin.style.display = 'none';
      lbl.textContent = '⚙ Generate & Download';
    }
  },

  // ── Settings modal ────────────────────────────────────────────────────────
  openSettings() {
    document.getElementById('settings-modal').classList.add('open');
  },
  closeSettings() {
    document.getElementById('settings-modal').classList.remove('open');
    // Re-apply default tech to field if it was changed
    const tech = Settings.get('defaultTech', CONFIG.defaultTechnician);
    if (tech) document.getElementById('f-tech').value = tech;
  },
  saveSetting(key, val) {
    Settings.set(key, val);
  },

  // ── Toast notifications ───────────────────────────────────────────────────
  _toastTimer: null,
  toast(msg, type = '') {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.className   = 'toast' + (type ? ' ' + type : '');
    el.style.display = 'block';
    clearTimeout(App._toastTimer);
    App._toastTimer = setTimeout(() => { el.style.display = 'none'; }, 5000);
  },
};

// ── Wire settings modal backdrop click ───────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('settings-modal')?.addEventListener('click', e => {
    if (e.target === e.currentTarget) App.closeSettings();
  });

  // Boot the app
  App.init();
});
