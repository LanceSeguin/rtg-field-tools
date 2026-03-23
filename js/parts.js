// ─────────────────────────────────────────────────────────────────────────────
// parts.js — Parts row state, rendering, and catalog lookup
//
// The parts catalog is loaded from parts_catalog.csv in the repo root.
// CSV format (must have a header row):
//   Label,PartNumber,Description
//   2230 Terminal Board,02230-3000-0001,"KIT, WIRE CONNECTIONS..."
//
// Label = shown in the dropdown list
// PartNumber = auto-fills the Part # field when selected
// Description = auto-fills the Description field when selected
//
// Edit parts_catalog.csv to add/remove/change parts. No code changes needed.
// ─────────────────────────────────────────────────────────────────────────────

const Parts = (() => {
  let _rows    = [];
  let _catalog = [];   // [{ part_number, description }, ...]

  // ── Load parts catalog from CSV ───────────────────────────────────────────
  async function loadCatalog() {
    try {
      const r = await fetch('./parts_catalog.csv?v=' + Date.now());
      if (!r.ok) {
        console.warn('parts_catalog.csv not found — parts dropdown will be empty');
        return;
      }
      const text = await r.text();
      _catalog = _parseCSV(text);
      _populateDropdown();
    } catch (e) {
      console.warn('Could not load parts catalog:', e.message);
    }
  }

  // CSV format: Label,PartNumber,Description
  // Label      = shown in the dropdown
  // PartNumber = fills the Part # field
  // Description = fills the Description field
  function _parseCSV(text) {
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
    if (lines.length < 2) return [];

    // Always skip header row
    const dataLines = lines.slice(1);

    return dataLines.map(line => {
      // Properly parse CSV with quoted fields that may contain commas
      const cols = [];
      let cur = '', inQuote = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') { inQuote = !inQuote; }
        else if (ch === ',' && !inQuote) { cols.push(cur.trim()); cur = ''; }
        else { cur += ch; }
      }
      cols.push(cur.trim());
      return {
        label:       cols[0] || '',
        part_number: cols[1] || '',
        description: cols[2] || '',
      };
    }).filter(p => p.label || p.part_number);
  }

  function _populateDropdown() {
    const sel = document.getElementById('pt-catalog');
    if (!sel) return;
    sel.innerHTML = '<option value="">— Select from catalog —</option>' +
      _catalog.map((p, i) =>
        `<option value="${i}">${p.label}</option>`
      ).join('');
  }

  // ── Catalog selection handler ──────────────────────────────────────────────
  function onCatalogSelect() {
    const sel = document.getElementById('pt-catalog');
    const idx = parseInt(sel?.value);
    if (isNaN(idx) || !_catalog[idx]) return;
    const part = _catalog[idx];
    const numEl  = document.getElementById('pt-num');
    const descEl = document.getElementById('pt-desc');
    if (numEl)  numEl.value  = part.part_number;
    if (descEl) descEl.value = part.description;
  }

  // ── Add a part row ────────────────────────────────────────────────────────
  function add() {
    const num  = document.getElementById('pt-num')?.value.trim()  || '';
    const desc = document.getElementById('pt-desc')?.value.trim() || '';
    const ser  = document.getElementById('pt-ser')?.value.trim()  || '';
    const qty  = parseFloat(document.getElementById('pt-qty')?.value) || 1;

    const row = { id: Date.now(), num, desc, ser, qty };
    _rows.push(row);
    _renderRow(row);

    // Clear inputs
    ['pt-num','pt-desc','pt-ser'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    const qtyEl = document.getElementById('pt-qty');
    if (qtyEl) qtyEl.value = '1';
    const catEl = document.getElementById('pt-catalog');
    if (catEl) catEl.value = '';
    document.getElementById('pt-num')?.focus();
  }

  function _renderRow(row) {
    const tr = document.createElement('tr');
    tr.id = 'pr-' + row.id;
    tr.innerHTML = `
      <td class="td-static">${row.num}</td>
      <td class="td-static">${row.desc}</td>
      <td class="td-static">${row.ser}</td>
      <td class="td-static">${row.qty}</td>
      <td><button class="btn-rm" onclick="Parts._rm(${row.id})" title="Remove">✕</button></td>`;
    document.getElementById('parts-body')?.appendChild(tr);
  }

  function _rm(id) {
    _rows = _rows.filter(r => r.id !== id);
    document.getElementById('pr-' + id)?.remove();
  }

  function getRows() { return _rows.map(r => ({ num: r.num, desc: r.desc, ser: r.ser, qty: r.qty })); }
  function count()   { return _rows.length; }

  return { loadCatalog, onCatalogSelect, add, _rm, getRows, count };
})();
