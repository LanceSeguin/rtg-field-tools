// ─────────────────────────────────────────────────────────────────────────────
// docx.js — Fills Work_Order_Master_Template.docx in the browser
//
// Uses ziplib.js (local, no CDN) to read/write the .docx ZIP format.
// Tokens split across XML runs are stitched at the paragraph level.
// ─────────────────────────────────────────────────────────────────────────────

const DOCX = (() => {

  const TEMPLATE_URL = './Work_Order_Master_Template.docx';

  // ── Fetch template ────────────────────────────────────────────────────────
  async function _fetchTemplate() {
    const r = await fetch(TEMPLATE_URL + '?v=' + Date.now());
    if (!r.ok) throw new Error(
      `Template fetch failed: ${r.status}. Make sure Work_Order_Master_Template.docx is in your GitHub repo root.`
    );
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
      .replace(/<img[^>]*>/gi, '')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .trim();
  }

  // ── Paragraph-level token replacement ────────────────────────────────────
  // Word splits tokens across <w:r> runs — stitch all <w:t> text in each
  // paragraph together, replace tokens, write back to first node.
  function _replaceParagraphTokens(paraXml, tokenMap) {
    const tRegex = /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g;
    const tNodes = [];
    let m;
    while ((m = tRegex.exec(paraXml)) !== null) {
      tNodes.push({ full: m[0], open: m[1], text: m[2], close: m[3], index: m.index });
    }
    if (!tNodes.length) return paraXml;

    const combined = tNodes.map(n => n.text).join('');
    let hasToken = false;
    for (const token of Object.keys(tokenMap)) {
      if (combined.includes(token)) { hasToken = true; break; }
    }
    if (!hasToken) return paraXml;

    let replaced = combined;
    for (const [token, value] of Object.entries(tokenMap)) {
      while (replaced.includes(token)) {
        replaced = replaced.split(token).join(_esc(value));
      }
    }

    // Write result into first <w:t>, blank the rest — iterate in reverse
    // so string indices stay valid
    let result = paraXml;
    for (let i = tNodes.length - 1; i >= 0; i--) {
      const node    = tNodes[i];
      const newText = i === 0 ? replaced : '';
      let openTag   = node.open;
      if (newText && /^ | $/.test(newText) && !openTag.includes('space')) {
        openTag = openTag.replace('>', ' xml:space="preserve">');
      }
      const newNode = openTag + newText + node.close;
      result = result.slice(0, node.index) + newNode + result.slice(node.index + node.full.length);
    }
    return result;
  }

  function _replaceAllTokens(xml, tokenMap) {
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
    return { rowXml: xml.slice(before, after + 7), start: before, end: after + 7 };
  }

  // ── Expand repeating rows ─────────────────────────────────────────────────
  function _expandRows(xml, tokenKey, rows, tokenMapFn, emptyTokens) {
    const tmpl = _findTemplateRow(xml, tokenKey);
    if (!tmpl) return xml;

    let newRows;
    if (rows && rows.length) {
      newRows = rows.map(r => _replaceAllTokens(tmpl.rowXml, tokenMapFn(r))).join('');
    } else {
      // No data — leave one blank row
      const emptyMap = {};
      emptyTokens.forEach(t => { emptyMap[`{{${t}}}`] = ''; });
      newRows = _replaceAllTokens(tmpl.rowXml, emptyMap);
    }
    return xml.slice(0, tmpl.start) + newRows + xml.slice(tmpl.end);
  }

  // ── Strip <w:sdt> wrappers, preserving inner <w:sdtContent> ─────────────────
  // Handles nested sdts correctly by tracking depth, not using regex.
  function _stripSdtWrappers(xml) {
    const out = [];
    let i = 0;
    while (i < xml.length) {
      if (xml.slice(i, i+7) === '<w:sdt>') {
        // Find matching </w:sdt> with depth tracking
        let depth = 1, j = i + 7;
        while (j < xml.length && depth > 0) {
          if      (xml.slice(j, j+7) === '<w:sdt>')  { depth++; j += 7; }
          else if (xml.slice(j, j+8) === '</w:sdt>') { depth--; if (depth > 0) j += 8; }
          else    j++;
        }
        // Extract content between first <w:sdtContent> and last </w:sdtContent>
        const block    = xml.slice(i, j + 8);
        const scStart  = block.indexOf('<w:sdtContent>');
        const scEnd    = block.lastIndexOf('</w:sdtContent>');
        if (scStart !== -1 && scEnd !== -1) {
          out.push(block.slice(scStart + 14, scEnd));
        } else {
          out.push(block);
        }
        i = j + 8;
      } else {
        out.push(xml[i++]);
      }
    }
    return out.join('');
  }

  // ── Main ──────────────────────────────────────────────────────────────────
  async function download(formData, filename) {

    // 1. Fetch template
    const buf = await _fetchTemplate();

    // 2. Parse ZIP
    const zipFiles = await ZipLib.readZip(buf);

    // 3. Get document.xml
    let xml = await ZipLib.getFileText(zipFiles, 'word/document.xml');

    // 4. Flat token map
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
      '{{customer_name}}':      formData.customerName   || '',
      '{{systemserial}}':       formData.systemSerial   || '',
    };

    // 5a. Strip elements that cause bracket/cursor indicators in Word:
    //   <w:sdt> content controls  → keep inner sdtContent, drop wrapper
    //   <w:permStart> / <w:permEnd> → editing permission markers, drop entirely
    xml = _stripSdtWrappers(xml);

    // 5b. Remove editing permission markers — these show as bracket cursors in Word
    xml = xml.replace(/<w:permStart[^/]*\/>/g, '');
    xml = xml.replace(/<w:permEnd[^/]*\/>/g, '');
    xml = xml.replace(/<w:permStart[^/]*\/>/g, '');
    xml = xml.replace(/<w:permEnd[^/]*\/>/g, '');

    // 5b. Replace flat tokens
    xml = _replaceAllTokens(xml, tokenMap);

    // 6. Expand labor rows
    const laborTokens = ['L.Date','L.DOW','L.Std.Reg','L.Std.OT','L.Std.Hol',
      'L.Next.Reg','L.Next.OT','L.Next.Hol','L.Second.Reg','L.Second.OT','L.Second.Hol','L.Notes'];
    xml = _expandRows(xml, 'L.Date', formData.laborRows || [], r => ({
      '{{L.Date}}':       r.date    || '',
      '{{L.DOW}}':        r.dow     || '',
      '{{L.Std.Reg}}':    r.std_reg || '',
      '{{L.Std.OT}}':     r.std_ot  || '',
      '{{L.Std.Hol}}':    r.std_hol || '',
      '{{L.Next.Reg}}':   r.nxt_reg || '',
      '{{L.Next.OT}}':    r.nxt_ot  || '',
      '{{L.Next.Hol}}':   r.nxt_hol || '',
      '{{L.Second.Reg}}': r.sec_reg || '',
      '{{L.Second.OT}}':  r.sec_ot  || '',
      '{{L.Second.Hol}}': r.sec_hol || '',
      '{{L.Notes}}':      r.notes   || '',
    }), laborTokens);

    // 7. Expand parts rows
    const partsTokens = ['P.Num','P.Desc','P.Serials','P.Qty'];
    xml = _expandRows(xml, 'P.Num', formData.partsRows || [], r => ({
      '{{P.Num}}':     r.num          || '',
      '{{P.Desc}}':    r.desc         || '',
      '{{P.Serials}}': r.ser          || '',
      '{{P.Qty}}':     String(r.qty || ''),
    }), partsTokens);

    // 8. Rebuild ALL zip files with modified document.xml
    const outFiles = {};
    for (const [name, entry] of Object.entries(zipFiles)) {
      if (name === 'word/document.xml') {
        outFiles[name] = xml;  // use our modified version
      } else {
        // Re-use original compressed data as-is by decompressing then rewriting
        outFiles[name] = await ZipLib.getFileAsBytes(zipFiles, name);
      }
    }

    // 9. Write new ZIP
    const bytes = await ZipLib.writeZip(outFiles);

    // 10. Download
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    const url = URL.createObjectURL(blob);
    const a   = Object.assign(document.createElement('a'), { href: url, download: filename });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  return { download };
})();
