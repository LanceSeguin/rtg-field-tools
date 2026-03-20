// ─────────────────────────────────────────────────────────────────────────────
// parts.js — Parts row state and rendering
// Edit this file to add fields to the parts table.
// ─────────────────────────────────────────────────────────────────────────────

const Parts = (() => {
  let _rows = [];

  // ── Public API ────────────────────────────────────────────────────────────

  /** Read input fields and add a part row */
  function add() {
    const num  = document.getElementById('pt-num').value.trim();
    const desc = document.getElementById('pt-desc').value.trim();
    const ser  = document.getElementById('pt-ser').value.trim();
    const qty  = parseFloat(document.getElementById('pt-qty').value) || 1;

    const row = { id: Date.now(), num, desc, ser, qty };
    _rows.push(row);
    _renderRow(row);

    // Clear inputs after add
    ['pt-num', 'pt-desc', 'pt-ser'].forEach(id => {
      document.getElementById(id).value = '';
    });
    document.getElementById('pt-qty').value = '1';
    document.getElementById('pt-num').focus();
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
    document.getElementById('parts-body').appendChild(tr);
  }

  function _rm(id) {
    _rows = _rows.filter(r => r.id !== id);
    document.getElementById('pr-' + id)?.remove();
  }

  /** Get all rows — used by docx.js at generation time */
  function getRows() {
    return _rows.map(r => ({
      num: r.num, desc: r.desc, ser: r.ser, qty: r.qty,
    }));
  }

  function count() { return _rows.length; }

  return { add, _rm, getRows, count };
})();
