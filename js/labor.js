// ─────────────────────────────────────────────────────────────────────────────
// labor.js — Labor row state and rendering
// Edit this file to change labor entry behavior or add/remove columns.
// ─────────────────────────────────────────────────────────────────────────────

const Labor = (() => {
  let _rows = [];

  const DAYS = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];

  // ── Date formatting ───────────────────────────────────────────────────────
  function _fmt(d) {
    return `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}/${d.getFullYear()}`;
  }

  // ── Row rendering ─────────────────────────────────────────────────────────
  // To add a column: add a field here, add a <th> in index.html, and add it to getRows()
  function _renderRow(row) {
    const inp = (field, width=52) =>
      `<input type="number" min="0" step="0.5" class="tbl-input" style="width:${width}px"
         value="${row[field]||''}"
         oninput="Labor._upd(${row.id},'${field}',this.value)">`;

    const tr = document.createElement('tr');
    tr.id = 'lr-' + row.id;
    tr.innerHTML = `
      <td><span class="td-static">${row.date}</span></td>
      <td><span class="td-dow">${row.dow}</span></td>
      <td>${inp('std_reg')}</td>
      <td>${inp('std_ot')}</td>
      <td>${inp('std_hol')}</td>
      <td>${inp('nxt_reg')}</td>
      <td>${inp('nxt_ot')}</td>
      <td>${inp('nxt_hol')}</td>
      <td>${inp('sec_reg')}</td>
      <td>${inp('sec_ot')}</td>
      <td>${inp('sec_hol')}</td>
      <td><input type="text" class="tbl-input" style="width:130px" placeholder="Notes"
           value="${row.notes||''}"
           oninput="Labor._upd(${row.id},'notes',this.value)"></td>
      <td><button class="btn-rm" onclick="Labor._rm(${row.id})" title="Remove row">✕</button></td>`;
    document.getElementById('labor-body').appendChild(tr);
  }

  // ── Public API ────────────────────────────────────────────────────────────

  /** Replace all rows with one row per date in the array (called from calendar import) */
  function populateFromDates(dates) {
    _rows = [];
    document.getElementById('labor-body').innerHTML = '';
    dates.forEach(d => _addRow(d));
  }

  /** Add a blank row for today (manual add button) */
  function addManual() {
    _addRow(new Date());
  }

  /** Internal: add a row for a specific Date object */
  function _addRow(dateObj) {
    const row = {
      id:      Date.now() + Math.random(),
      date:    _fmt(dateObj),
      dow:     DAYS[dateObj.getDay()],
      std_reg: '', std_ot: '', std_hol: '',
      nxt_reg: '', nxt_ot: '', nxt_hol: '',
      sec_reg: '', sec_ot: '', sec_hol: '',
      notes:   '',
    };
    _rows.push(row);
    _renderRow(row);
  }

  /** Update a field value when user types in a cell */
  function _upd(id, field, value) {
    const row = _rows.find(r => r.id === id);
    if (row) row[field] = value;
  }

  /** Remove a row */
  function _rm(id) {
    _rows = _rows.filter(r => r.id !== id);
    document.getElementById('lr-' + id)?.remove();
  }

  /** Get all rows — used by docx.js at generation time */
  function getRows() {
    return _rows.map(r => ({
      date:    r.date,
      dow:     r.dow,
      std_reg: r.std_reg || '',   std_ot: r.std_ot || '',   std_hol: r.std_hol || '',
      nxt_reg: r.nxt_reg || '',   nxt_ot: r.nxt_ot || '',   nxt_hol: r.nxt_hol || '',
      sec_reg: r.sec_reg || '',   sec_ot: r.sec_ot || '',   sec_hol: r.sec_hol || '',
      notes:   r.notes   || '',
    }));
  }

  function count() { return _rows.length; }

  return { populateFromDates, addManual, _upd, _rm, getRows, count };
})();
