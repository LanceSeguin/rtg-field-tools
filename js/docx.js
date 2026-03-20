// ─────────────────────────────────────────────────────────────────────────────
// docx.js — .docx generation in pure JavaScript (no libraries)
//
// Builds a valid Open XML .docx file and triggers a browser download.
// Edit this file to change the layout or content of the generated document.
// ─────────────────────────────────────────────────────────────────────────────

const DOCX = (() => {

  // ══════════════════════════════════════════════════════════════════════════
  // [1] ZIP builder (store-only, no compression — keeps code simple)
  // ══════════════════════════════════════════════════════════════════════════

  function _u8(str)  { return new TextEncoder().encode(str); }
  function _u16(n)   { return [n & 0xFF, (n >> 8) & 0xFF]; }
  function _u32(n)   { return [n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >> 24) & 0xFF]; }

  function _crc32(data) {
    const table = _crc32._t || (_crc32._t = Array.from({length:256}, (_,i) => {
      let n = i;
      for (let j=0;j<8;j++) n = (n&1) ? (0xEDB88320^(n>>>1)) : (n>>>1);
      return n;
    }));
    let c = 0xFFFFFFFF;
    for (const b of data) c = table[(c^b)&0xFF] ^ (c>>>8);
    return (c^0xFFFFFFFF) >>> 0;
  }

  function _buildZip(files) {
    const locals = [], centrals = [];
    let offset = 0;

    for (const {name, data} of files) {
      const nameB = _u8(name);
      const bytes = typeof data === 'string' ? _u8(data) : data;
      const crc   = _crc32(bytes);
      const local = Uint8Array.from([
        0x50,0x4B,0x03,0x04, 20,0, 0,0, 0,0, 0,0, 0,0,
        ..._u32(crc), ..._u32(bytes.length), ..._u32(bytes.length),
        ..._u16(nameB.length), 0,0,
        ...nameB, ...bytes,
      ]);
      const central = [...[
        0x50,0x4B,0x01,0x02, 20,0, 20,0, 0,0, 0,0, 0,0, 0,0,
        ..._u32(crc), ..._u32(bytes.length), ..._u32(bytes.length),
        ..._u16(nameB.length), 0,0, 0,0, 0,0, 0,0, 0,0,
        ..._u32(0), // offset patched below
        ...nameB,
      ]];
      // Patch offset at byte 42
      _u32(offset).forEach((b,i) => central[42+i] = b);
      locals.push(local); centrals.push(central);
      offset += local.length;
    }

    const centralStart = offset;
    const centralBytes = Uint8Array.from(centrals.flat());
    const eocd = Uint8Array.from([
      0x50,0x4B,0x05,0x06, 0,0, 0,0,
      ..._u16(files.length), ..._u16(files.length),
      ..._u32(centralBytes.length), ..._u32(centralStart), 0,0,
    ]);

    const total = locals.reduce((a,l) => a+l.length, 0) + centralBytes.length + eocd.length;
    const out   = new Uint8Array(total);
    let pos = 0;
    locals.forEach(l => { out.set(l, pos); pos += l.length; });
    out.set(centralBytes, pos); pos += centralBytes.length;
    out.set(eocd, pos);
    return out;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // [2] XML helpers
  // ══════════════════════════════════════════════════════════════════════════

  function _esc(s) {
    return String(s || '')
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  /** Single text run with optional formatting */
  function wRun(text, {bold, italic, underline, size, color} = {}) {
    if (!text && text !== 0) return '';
    const rpr = [
      bold      ? '<w:b/>'                                              : '',
      italic    ? '<w:i/>'                                              : '',
      underline ? '<w:u w:val="single"/>'                               : '',
      size      ? `<w:sz w:val="${size*2}"/><w:szCs w:val="${size*2}"/>`: '',
      color     ? `<w:color w:val="${String(color).replace('#','')}"/>` : '',
    ].join('');
    const rprXml = rpr ? `<w:rPr>${rpr}</w:rPr>` : '';
    return `<w:r>${rprXml}<w:t xml:space="preserve">${_esc(text)}</w:t></w:r>`;
  }

  /** Paragraph wrapping runs, optional style and alignment */
  function wPara(runs, style, align) {
    const ppr = [
      style ? `<w:pStyle w:val="${style}"/>` : '',
      align ? `<w:jc w:val="${align}"/>`     : '',
    ].join('');
    return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${runs}</w:p>`;
  }

  /** Table cell */
  function wCell(content, {bold=false, shade, width, span}={}) {
    const tcpr = [
      width ? `<w:tcW w:w="${width}" w:type="dxa"/>`                    : '',
      span  ? `<w:gridSpan w:val="${span}"/>`                           : '',
      shade ? `<w:shd w:val="clear" w:color="auto" w:fill="${shade}"/>` : '',
      `<w:tcMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>`,
    ].join('');
    const runs = typeof content === 'string' ? wRun(content, {bold}) : content;
    return `<w:tc><w:tcPr>${tcpr}</w:tcPr>${wPara(runs)}</w:tc>`;
  }

  /** Table row */
  function wRow(cells) { return `<w:tr>${cells}</w:tr>`; }

  /** Full-width table with header row (dark) + data rows */
  function wTable(headerCells, dataRows, totalWidth=9360) {
    const hdr = wRow(headerCells.map(h =>
      wCell(h, {bold:true, shade:'1F4E79', width: Math.floor(totalWidth/headerCells.length)})
    ).join(''));
    const body = dataRows.map(cells =>
      wRow(cells.map(c => wCell(c)).join(''))
    ).join('');
    return `<w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="${totalWidth}" w:type="dxa"/>
      </w:tblPr>
      ${hdr}${body}
    </w:tbl>`;
  }

  /** Section heading paragraph */
  function wHeading(text) {
    return wPara(wRun(text, {bold:true, underline:true, size:12}));
  }

  const EMPTY_PARA = '<w:p/>';

  // ══════════════════════════════════════════════════════════════════════════
  // [3] Rich text editor → Word XML
  // ══════════════════════════════════════════════════════════════════════════

  async function _rteToXml(editorEl, imgFiles) {
    let xml = '';
    for (const node of editorEl.childNodes) {
      xml += await _nodeToXml(node, imgFiles);
    }
    return xml || EMPTY_PARA;
  }

  async function _nodeToXml(node, imgFiles) {
    if (node.nodeType === Node.TEXT_NODE) {
      const t = node.textContent;
      return t ? wPara(wRun(t)) : '';
    }
    const tag = node.nodeName;
    if (tag === 'IMG')  return await _imgToXml(node, imgFiles);
    if (tag === 'BR')   return EMPTY_PARA;

    if (tag === 'P' || tag === 'DIV') {
      const runs = await _inlineToRuns(node, imgFiles);
      if (tag === 'P' && node.style?.textAlign === 'center')
        return wPara(runs, '', 'center');
      return wPara(runs);
    }
    if (tag === 'UL' || tag === 'OL') return await _listToXml(node, tag, imgFiles);
    if (tag === 'LI')  return await _listItemToXml(node, tag, imgFiles);

    // Generic block fallback
    const runs = await _inlineToRuns(node, imgFiles);
    return runs ? wPara(runs) : '';
  }

  async function _inlineToRuns(el, imgFiles) {
    let runs = '';
    for (const child of el.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        runs += wRun(child.textContent);
      } else if (child.nodeName === 'IMG') {
        runs += await _imgToXml(child, imgFiles);
      } else if (child.nodeName === 'BR') {
        runs += '<w:r><w:br/></w:r>';
      } else {
        const tag  = child.nodeName;
        const bold = tag === 'B' || tag === 'STRONG' || child.style?.fontWeight === 'bold';
        const ital = tag === 'I' || tag === 'EM'    || child.style?.fontStyle   === 'italic';
        const und  = tag === 'U';
        // Recurse for nested inline elements
        for (const gc of child.childNodes) {
          if (gc.nodeType === Node.TEXT_NODE)
            runs += wRun(gc.textContent, {bold, italic:ital, underline:und});
          else if (gc.nodeName === 'IMG')
            runs += await _imgToXml(gc, imgFiles);
          else
            runs += wRun(gc.textContent || '', {bold, italic:ital, underline:und});
        }
        if (!child.childNodes.length)
          runs += wRun(child.textContent, {bold, italic:ital, underline:und});
      }
    }
    return runs;
  }

  async function _listToXml(listEl, tag, imgFiles) {
    let xml = '';
    for (const li of listEl.querySelectorAll('li')) {
      const runs  = await _inlineToRuns(li, imgFiles);
      const numId = tag === 'OL' ? 2 : 1;
      xml += `<w:p>
        <w:pPr><w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="${numId}"/>
        </w:numPr></w:pPr>
        ${runs}
      </w:p>`;
    }
    return xml;
  }

  async function _listItemToXml(li, tag, imgFiles) {
    const runs = await _inlineToRuns(li, imgFiles);
    return `<w:p>
      <w:pPr><w:numPr>
        <w:ilvl w:val="0"/><w:numId w:val="1"/>
      </w:numPr></w:pPr>${runs}
    </w:p>`;
  }

  async function _imgToXml(imgEl, imgFiles) {
    const src = imgEl.src || imgEl.getAttribute('src') || '';
    if (!src || !src.startsWith('data:')) return '';

    const rId  = `rId${imgFiles.length + 20}`;
    const ext  = (src.match(/data:image\/(\w+)/) || [,'png'])[1].replace('jpeg','jpg');
    const b64  = src.split(',')[1];
    const bin  = atob(b64);
    const bytes= new Uint8Array(bin.length);
    for (let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);

    // Determine display size
    const maxIn = 6.0;
    const dispW = imgEl.naturalWidth  || imgEl.width  || 400;
    const dispH = imgEl.naturalHeight || imgEl.height || 300;
    const wIn   = Math.min(dispW / 96, maxIn);
    const ar    = dispH / Math.max(dispW, 1);
    const wEmu  = Math.round(wIn * 914400);
    const hEmu  = Math.round(wIn * ar * 914400);

    imgFiles.push({ rId, ext, bytes });

    return `<w:p><w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="${wEmu}" cy="${hEmu}"/>
        <wp:docPr id="${imgFiles.length}" name="Image${imgFiles.length}"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill>
                <a:blip r:embed="${rId}"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="${wEmu}" cy="${hEmu}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // [4] Build the complete .docx
  //
  // Edit the section between "DOCUMENT LAYOUT" comments to change the
  // structure / order of sections in the output document.
  // ══════════════════════════════════════════════════════════════════════════

  async function build(data) {
    const imgFiles  = [];
    const summXml   = await _rteToXml(document.getElementById('rte-summary'),  imgFiles);
    const followXml = await _rteToXml(document.getElementById('rte-followup'), imgFiles);

    // ── Image relationships ──────────────────────────────────────────────────
    const imgRels = imgFiles.map(f =>
      `<Relationship Id="${f.rId}"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        Target="media/${f.rId}.${f.ext}"/>`
    ).join('');

    // ── Numbering (bullet + ordered lists) ──────────────────────────────────
    const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

    // ── Labor table ──────────────────────────────────────────────────────────
    const laborRows = data.laborRows || [];
    const laborXml  = laborRows.length ? `
      ${wHeading('LABOR')}${EMPTY_PARA}
      ${wTable(
        ['Date','Day','Std Reg','Std OT','Std Hol','Nxt Reg','Nxt OT','Nxt Hol','2nd Reg','2nd OT','2nd Hol','Notes'],
        laborRows.map(r => [r.date,r.dow,r.std_reg,r.std_ot,r.std_hol,r.nxt_reg,r.nxt_ot,r.nxt_hol,r.sec_reg,r.sec_ot,r.sec_hol,r.notes])
      )}${EMPTY_PARA}` : '';

    // ── Parts table ──────────────────────────────────────────────────────────
    const partsRows = data.partsRows || [];
    const partsXml  = partsRows.length ? `
      ${wHeading('PARTS')}${EMPTY_PARA}
      ${wTable(
        ['Part #', 'Description', 'Serial(s)', 'Qty'],
        partsRows.map(r => [r.num, r.desc, r.ser, String(r.qty)])
      )}${EMPTY_PARA}` : '';

    // ════════════════════════════════════════════════════════════════════════
    // DOCUMENT LAYOUT — edit here to change section order or add new sections
    // ════════════════════════════════════════════════════════════════════════
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<w:body>

${wPara(wRun('SERVICE WORK ORDER', {bold:true, underline:true, size:14}), '', 'center')}
${EMPTY_PARA}

${wTable(['CUSTOMER','SERVICE AGENCY'], [
  [_esc(data.customerName || ''), _esc(data.techName || '')],
  ['Purchase Order: ' + _esc(data.poNumber || ''), 'Agency Order: ' + _esc(data.serviceOrder || '')],
  ['', 'Product Line: ' + _esc(data.productLine || '')],
  ['', 'System/Serial: ' + _esc(data.systemSerial || '')],
])}
${EMPTY_PARA}

${wTable(['SERVICE ADDRESS','LOCATION CONTACT'], [
  [_esc(data.serviceAddress || ''), 'Name: '  + _esc(data.contactName  || '')],
  ['', 'Phone: ' + _esc(data.contactPhone || '')],
  ['', 'Email: ' + _esc(data.contactEmail || '')],
])}
${EMPTY_PARA}

${wTable(['SCOPE OF WORK'], [[_esc(data.scope || '')]])}
${EMPTY_PARA}

${wHeading('SERVICE REPORT')}${EMPTY_PARA}

<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="9360" w:type="dxa"/>
  </w:tblPr>
  ${wRow(wCell('SUMMARY', {bold:true, shade:'1F4E79', width:9360}))}
  <w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${summXml}</w:tc></w:tr>
</w:tbl>
${EMPTY_PARA}

<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="9360" w:type="dxa"/>
  </w:tblPr>
  ${wRow(wCell('REQUIRED FOLLOW UP (IF APPLICABLE)', {bold:true, shade:'1F4E79', width:9360}))}
  <w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${followXml}</w:tc></w:tr>
</w:tbl>
${EMPTY_PARA}

${laborXml}
${partsXml}

${wTable(['SERVICE TECHNICIAN SIGNATURE','CUSTOMER SIGNATURE / PRINTED NAME'], [
  ['', ''],
  ['Date: ____________________', 'Date: ____________________'],
])}

<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/>
</w:sectPr>

</w:body>
</w:document>`;

    // ── Assemble zip files ────────────────────────────────────────────────
    const files = [
      { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Default Extension="jpg"  ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="gif"  ContentType="image/gif"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/numbering.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>` },

      { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>` },

      { name: 'word/_rels/document.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    Target="numbering.xml"/>
  ${imgRels}
</Relationships>` },

      { name: 'word/document.xml',  data: docXml },
      { name: 'word/numbering.xml', data: numberingXml },

      { name: 'word/styles.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:tblPr><w:tblBorders>
      <w:top    w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left   w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right  w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders></w:tblPr>
  </w:style>
</w:styles>` },
    ];

    // Add image binary files
    imgFiles.forEach(f => {
      files.push({ name: `word/media/${f.rId}.${f.ext}`, data: f.bytes });
    });

    return _buildZip(files);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // [5] Download trigger
  // ══════════════════════════════════════════════════════════════════════════

  async function download(data, filename) {
    const bytes = await build(data);
    const blob  = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    const url = URL.createObjectURL(blob);
    const a   = Object.assign(document.createElement('a'), { href: url, download: filename });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  return { build, download };
})();
