// ─────────────────────────────────────────────────────────────────────────────
// docx.js — Fills Work_Order_Master_Template.docx in the browser
//
// KEY PROBLEM SOLVED: Word splits tokens across multiple <w:r><w:t> runs
// when you edit a document (e.g. {{PON</w:t>...<w:t>umber}}). A naive
// string replace misses these. We fix this by stitching all <w:t> text
// nodes within each <w:p> paragraph together, doing the replacement on the
// combined string, then writing back — exactly what tokens.py did in Python.
// ─────────────────────────────────────────────────────────────────────────────

const DOCX = (() => {

  const TEMPLATE_URL = './Work_Order_Master_Template.docx';

  // ── Fetch template ────────────────────────────────────────────────────────
  async function _fetchTemplate() {
    const r = await fetch(TEMPLATE_URL + '?v=' + Date.now());
    if (!r.ok) throw new Error(
      `Template fetch failed: ${r.status}. Make sure Work_Order_Master_Template.docx is in the root of your GitHub repo.`
    );
    return r.arrayBuffer();
  }

  // ── XML helpers ───────────────────────────────────────────────────────────
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
      .replace(/<img[^>]*>/gi, '')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .trim();
  }

  // ── CORE FIX: stitch split runs then replace token in a paragraph ─────────
  // Word splits tokens across <w:r> runs when the document is edited.
  // e.g. {{PON</w:t></w:r>...<w:r><w:t>umber}} appears as two separate runs.
  // We collect ALL <w:t> text in the paragraph, do the replace on the joined
  // string, then put the result back into the FIRST <w:t> and blank the rest.
  function _replaceParagraphTokens(paraXml, tokenMap) {
    // Extract all <w:t> nodes and their text
    const tNodes = [];
    const tRegex = /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g;
    let m;
    while ((m = tRegex.exec(paraXml)) !== null) {
      tNodes.push({ full: m[0], open: m[1], text: m[2], close: m[3], index: m.index });
    }
    if (!tNodes.length) return paraXml;

    // Join all text to find tokens that span runs
    const combined = tNodes.map(n => n.text).join('');

    // Check if any token exists in combined text
    let hasToken = false;
    for (const token of Object.keys(tokenMap)) {
      if (combined.includes(token)) { hasToken = true; break; }
    }
    if (!hasToken) return paraXml;

    // Replace all tokens in combined text
    let replaced = combined;
    for (const [token, value] of Object.entries(tokenMap)) {
      // Use split/join for reliable replacement (no regex special char issues)
      while (replaced.includes(token)) {
        replaced = replaced.split(token).join(_esc(value));
      }
    }

    // Write replaced text into first <w:t>, blank out the rest
    let result = paraXml;
    // Replace each original <w:t> node — first gets full value, rest get empty
    for (let i = tNodes.length - 1; i >= 0; i--) {
      const node  = tNodes[i];
      const newText = i === 0 ? replaced : '';
      // Preserve xml:space="preserve" if present, otherwise add it for safety
      let openTag = node.open;
      if (newText.match(/^ | $/) && !openTag.includes('space')) {
        openTag = openTag.replace('>', ' xml:space="preserve">');
      }
      const newNode = openTag + newText + node.close;
      result = result.slice(0, node.index) + newNode + result.slice(node.index + node.full.length);
    }
    return result;
  }

  // ── Apply token replacement across entire XML ─────────────────────────────
  function _replaceAllTokens(xml, tokenMap) {
    // Process paragraph by paragraph
    return xml.replace(/(<w:p[ >][\s\S]*?<\/w:p>)/g, para =>
      _replaceParagraphTokens(para, tokenMap)
    );
  }

  // ── Find template row containing a token ─────────────────────────────────
  function _findTemplateRow(xml, token) {
    const pos = xml.indexOf(`{{${token}}}`);
    if (pos === -1) return null;
    const before = Math.max(
      xml.lastIndexOf('<w:tr ', pos),
      xml.lastIndexOf('<w:tr>', pos)
    );
    if (before === -1) return null;
    const after = xml.indexOf('</w:tr>', pos);
    if (after === -1) return null;
    return {
      rowXml: xml.slice(before, after + 7),
      start:  before,
      end:    after + 7,
    };
  }

  // ── Expand a repeating table section ─────────────────────────────────────
  function _expandRows(xml, tokenKey, rows, tokenMapFn) {
    const tmpl = _findTemplateRow(xml, tokenKey);
    if (!tmpl) return xml;

    let newRows = '';
    if (rows && rows.length) {
      newRows = rows.map(r => {
        // Replace tokens in the template row for this data row
        return _replaceAllTokens(tmpl.rowXml, tokenMapFn(r));
      }).join('');
    }
    // If no rows, leave one empty row so the table doesn't break
    if (!newRows) {
      newRows = _replaceAllTokens(tmpl.rowXml,
        Object.fromEntries(
          ['L.Date','L.DOW','L.Std.Reg','L.Std.OT','L.Std.Hol',
           'L.Next.Reg','L.Next.OT','L.Next.Hol',
           'L.Second.Reg','L.Second.OT','L.Second.Hol','L.Notes',
           'P.Num','P.Desc','P.Serials','P.Qty']
          .map(t => [`{{${t}}}`, ''])
        )
      );
    }
    return xml.slice(0, tmpl.start) + newRows + xml.slice(tmpl.end);
  }

  // ── Main: fetch → fill → download ────────────────────────────────────────
  async function download(formData, filename) {
    if (typeof PizZip === 'undefined') {
      throw new Error('PizZip not loaded — check index.html script tags');
    }

    // 1. Fetch template
    const buf = await _fetchTemplate();

    // 2. Unzip
    const zip = new PizZip(buf);
    const docFile = zip.file('word/document.xml');
    if (!docFile) throw new Error('Template is missing word/document.xml');
    let xml = docFile.asText();

    // 3. Build flat token map
    const tokenMap = {
      '{{CustomerName}}':       formData.customerName   || '',
      '{{PONumber}}':           formData.poNumber       || '',
      '{{ServiceAgency}}':      formData.serviceAgency  || '',
      '{{ServiceTechnician}}':  formData.techName       || '',
      '{{ServiceAgencyOrder}}': formData.serviceOrder   || '',
      '{{ProductLine}}':        formData.productLine    || '',
      '{{SystemSerial}}':       formData.systemSerial   || '',
      '{{ServiceAddress}}':     formData.serviceAddress || '',
      '{{ContactName}}':        formData.contactName    || '',
      '{{ContactNumber}}':      formData.contactPhone   || '',
      '{{ContactEmail}}':       formData.contactEmail   || '',
      '{{Scope}}':              formData.scope          || '',
      '{{Summary}}':            _rteText('rte-summary'),
      '{{FollowUp}}':           _rteText('rte-followup'),
      // Legacy lowercase aliases
      '{{customer_name}}':      formData.customerName   || '',
      '{{systemserial}}':       formData.systemSerial   || '',
    };

    // 4. Replace all flat tokens (handles split runs)
    xml = _replaceAllTokens(xml, tokenMap);

    // 5. Expand labor rows
    xml = _expandRows(xml, 'L.Date', formData.laborRows || [], r => ({
      '{{L.Date}}':        r.date    || '',
      '{{L.DOW}}':         r.dow     || '',
      '{{L.Std.Reg}}':     r.std_reg || '',
      '{{L.Std.OT}}':      r.std_ot  || '',
      '{{L.Std.Hol}}':     r.std_hol || '',
      '{{L.Next.Reg}}':    r.nxt_reg || '',
      '{{L.Next.OT}}':     r.nxt_ot  || '',
      '{{L.Next.Hol}}':    r.nxt_hol || '',
      '{{L.Second.Reg}}':  r.sec_reg || '',
      '{{L.Second.OT}}':   r.sec_ot  || '',
      '{{L.Second.Hol}}':  r.sec_hol || '',
      '{{L.Notes}}':       r.notes   || '',
    }));

    // 6. Expand parts rows
    xml = _expandRows(xml, 'P.Num', formData.partsRows || [], r => ({
      '{{P.Num}}':     r.num          || '',
      '{{P.Desc}}':    r.desc         || '',
      '{{P.Serials}}': r.ser          || '',
      '{{P.Qty}}':     String(r.qty || ''),
    }));

    // 7. Write back and generate
    zip.file('word/document.xml', xml);
    const blob = zip.generate({
      type:        'blob',
      mimeType:    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
    });

    // 8. Download
    const url = URL.createObjectURL(blob);
    const a   = Object.assign(document.createElement('a'), { href: url, download: filename });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  return { download };
})();
