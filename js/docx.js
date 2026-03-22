// ─────────────────────────────────────────────────────────────────────────────
// docx.js — Fills Work_Order_Master_Template.docx in the browser
//
// Strategy:
//   1. Fetch the template .docx from GitHub Pages
//   2. Unzip it with PizZip
//   3. Replace flat tokens ({{CustomerName}} etc.) with form values
//   4. For labor rows: find the template row containing {{L.Date}},
//      clone it once per labor entry, replace tokens, remove template row
//   5. Same for parts rows with {{P.Num}}
//   6. Re-zip and download
//
// This mirrors exactly what the original Python tokens.py + doc_table.py did.
// ─────────────────────────────────────────────────────────────────────────────

const DOCX = (() => {

  const TEMPLATE_URL = './Work_Order_Master_Template.docx';

  // ── Fetch template ────────────────────────────────────────────────────────
  async function _fetchTemplate() {
    const r = await fetch(TEMPLATE_URL + '?v=' + Date.now());
    if (!r.ok) throw new Error(`Template fetch failed: ${r.status} — make sure Work_Order_Master_Template.docx is in your GitHub repo root`);
    return r.arrayBuffer();
  }

  // ── XML escape ────────────────────────────────────────────────────────────
  function _esc(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ── Extract plain text from rich text editor ──────────────────────────────
  function _rteText(editorId) {
    const el = document.getElementById(editorId);
    if (!el) return '';
    return el.innerHTML
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n')
      .replace(/<\/div>/gi, '\n')
      .replace(/<img[^>]*>/gi, '[image]')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .trim();
  }

  // ── Replace a flat token throughout XML ───────────────────────────────────
  // Tokens can be split across multiple <w:r><w:t> runs by Word.
  // We first stitch all text nodes in each paragraph together, do the
  // replacement on the combined string, then write back to the first run.
  function _replaceToken(xml, token, value) {
    const escaped = _esc(value);
    // Simple case: token appears intact in a single <w:t> node
    const simple = xml.split(`{{${token}}}`).join(escaped);
    if (simple !== xml) return simple;
    // Regex to catch token split by XML attributes/namespace noise
    const inner = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
                       .replace(/\s+/g, '[\\s\\S]{0,50}');
    const rx = new RegExp(`\\{\\{[\\s]*${inner}[\\s]*\\}\\}`, 'g');
    return xml.replace(rx, escaped);
  }

  // ── Replace all flat tokens in XML ───────────────────────────────────────
  function _replaceFlatTokens(xml, data) {
    const flat = {
      'CustomerName':       data.customerName,
      'PONumber':           data.poNumber,
      'ServiceAgency':      data.serviceAgency,
      'ServiceTechnician':  data.techName,
      'ServiceAgencyOrder': data.serviceOrder,
      'ProductLine':        data.productLine,
      'SystemSerial':       data.systemSerial,
      'ServiceAddress':     data.serviceAddress,
      'ContactName':        data.contactName,
      'ContactNumber':      data.contactPhone,
      'ContactEmail':       data.contactEmail,
      'Scope':              data.scope,
      'Summary':            data.summary,
      'FollowUp':           data.followUp,
      // Legacy lowercase variants
      'customer_name':      data.customerName,
      'systemserial':       data.systemSerial,
    };
    let out = xml;
    for (const [token, value] of Object.entries(flat)) {
      out = _replaceToken(out, token, value || '');
    }
    return out;
  }

  // ── Find a table row (<w:tr>...</w:tr>) containing a given token ──────────
  function _findTemplateRow(xml, token) {
    const start = xml.indexOf(`{{${token}}}`);
    if (start === -1) return null;
    // Walk backwards to find opening <w:tr>
    const before  = xml.lastIndexOf('<w:tr ', start);
    const before2 = xml.lastIndexOf('<w:tr>', start);
    const rowStart = Math.max(before, before2);
    if (rowStart === -1) return null;
    // Walk forwards to find matching </w:tr>
    const rowEnd = xml.indexOf('</w:tr>', start);
    if (rowEnd === -1) return null;
    return {
      rowXml: xml.slice(rowStart, rowEnd + '</w:tr>'.length),
      start:  rowStart,
      end:    rowEnd + '</w:tr>'.length,
    };
  }

  // ── Expand labor rows ─────────────────────────────────────────────────────
  function _expandLaborRows(xml, laborRows) {
    const tmpl = _findTemplateRow(xml, 'L.Date');
    if (!tmpl) return xml;  // no template row found — leave as-is

    if (!laborRows || !laborRows.length) {
      // No labor data — replace template row with an empty row
      const emptyRow = tmpl.rowXml
        .replace(/\{\{L\.[^}]+\}\}/g, '');
      return xml.slice(0, tmpl.start) + emptyRow + xml.slice(tmpl.end);
    }

    const newRows = laborRows.map(r => {
      let row = tmpl.rowXml;
      const map = {
        'L.Date':        r.date       || '',
        'L.DOW':         r.dow        || '',
        'L.Std.Reg':     r.std_reg    || '',
        'L.Std.OT':      r.std_ot     || '',
        'L.Std.Hol':     r.std_hol    || '',
        'L.Next.Reg':    r.nxt_reg    || '',
        'L.Next.OT':     r.nxt_ot     || '',
        'L.Next.Hol':    r.nxt_hol    || '',
        'L.Second.Reg':  r.sec_reg    || '',
        'L.Second.OT':   r.sec_ot     || '',
        'L.Second.Hol':  r.sec_hol    || '',
        'L.Notes':       r.notes      || '',
      };
      for (const [token, value] of Object.entries(map)) {
        row = row.split(`{{${token}}}`).join(_esc(value));
      }
      return row;
    }).join('');

    return xml.slice(0, tmpl.start) + newRows + xml.slice(tmpl.end);
  }

  // ── Expand parts rows ─────────────────────────────────────────────────────
  function _expandPartsRows(xml, partsRows) {
    const tmpl = _findTemplateRow(xml, 'P.Num');
    if (!tmpl) return xml;

    if (!partsRows || !partsRows.length) {
      const emptyRow = tmpl.rowXml.replace(/\{\{P\.[^}]+\}\}/g, '');
      return xml.slice(0, tmpl.start) + emptyRow + xml.slice(tmpl.end);
    }

    const newRows = partsRows.map(r => {
      let row = tmpl.rowXml;
      const map = {
        'P.Num':     r.num  || '',
        'P.Desc':    r.desc || '',
        'P.Serials': r.ser  || '',
        'P.Qty':     String(r.qty || ''),
      };
      for (const [token, value] of Object.entries(map)) {
        row = row.split(`{{${token}}}`).join(_esc(value));
      }
      return row;
    }).join('');

    return xml.slice(0, tmpl.start) + newRows + xml.slice(tmpl.end);
  }

  // ── Main: fetch, fill, download ───────────────────────────────────────────
  async function download(formData, filename) {
    if (typeof PizZip === 'undefined') {
      throw new Error('PizZip not loaded — check that CDN scripts are in index.html');
    }

    // 1. Fetch template
    const buf = await _fetchTemplate();

    // 2. Unzip
    const zip = new PizZip(buf);

    // 3. Get document.xml
    const docFile = zip.file('word/document.xml');
    if (!docFile) throw new Error('word/document.xml not found in template');
    let xml = docFile.asText();

    // 4. Build data object
    const data = {
      customerName:   formData.customerName   || '',
      poNumber:       formData.poNumber       || '',
      serviceAgency:  formData.serviceAgency  || '',
      techName:       formData.techName       || '',
      serviceOrder:   formData.serviceOrder   || '',
      productLine:    formData.productLine    || '',
      systemSerial:   formData.systemSerial   || '',
      serviceAddress: formData.serviceAddress || '',
      contactName:    formData.contactName    || '',
      contactPhone:   formData.contactPhone   || '',
      contactEmail:   formData.contactEmail   || '',
      scope:          formData.scope          || '',
      summary:        _rteText('rte-summary'),
      followUp:       _rteText('rte-followup'),
    };

    // 5. Replace flat tokens
    xml = _replaceFlatTokens(xml, data);

    // 6. Expand labor rows
    xml = _expandLaborRows(xml, formData.laborRows || []);

    // 7. Expand parts rows
    xml = _expandPartsRows(xml, formData.partsRows || []);

    // 8. Write back
    zip.file('word/document.xml', xml);

    // 9. Generate and download
    const blob = zip.generate({
      type:        'blob',
      mimeType:    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
    });

    const url = URL.createObjectURL(blob);
    const a   = Object.assign(document.createElement('a'), {
      href:     url,
      download: filename,
    });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  return { download };
})();
