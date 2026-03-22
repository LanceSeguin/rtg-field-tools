// ─────────────────────────────────────────────────────────────────────────────
// docx.js — Fills Work_Order_Master_Template.docx in the browser
// Uses ziplib.js (local) to read/write the .docx ZIP format.
// ─────────────────────────────────────────────────────────────────────────────

const DOCX = (() => {

  const TEMPLATE_URL = './Work_Order_Master_Template.docx';

  async function _fetchTemplate() {
    const r = await fetch(TEMPLATE_URL + '?v=' + Date.now());
    if (!r.ok) throw new Error(`Template fetch failed: ${r.status}. Make sure Work_Order_Master_Template.docx is in your GitHub repo root.`);
    return r.arrayBuffer();
  }

  function _esc(s) {
    return String(s || '')
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function _rteText(editorId) {
    const el = document.getElementById(editorId);
    if (!el) return '';
    return el.innerHTML
      .replace(/<br\s*\/?>/gi,'\n').replace(/<\/p>/gi,'\n')
      .replace(/<\/div>/gi,'\n').replace(/<img[^>]*>/gi,'')
      .replace(/<[^>]+>/g,'').replace(/&nbsp;/g,' ')
      .replace(/&amp;/g,'&').replace(/&lt;/g,'<')
      .replace(/&gt;/g,'>').replace(/&quot;/g,'"')
      .trim();
  }

  function _dataUrlToBytes(dataUrl) {
    const b64 = dataUrl.split(',')[1];
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i=0;i<bin.length;i++) out[i]=bin.charCodeAt(i);
    return out;
  }

  function _mimeExt(dataUrl) {
    const m = dataUrl.match(/data:image\/(\w+)/);
    return (m ? m[1] : 'png').replace('jpeg','jpg');
  }

  function _makeImageXml(wEmu, hEmu, rId) {
    return `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${wEmu}" cy="${hEmu}"/><wp:docPr id="1" name="img"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${wEmu}" cy="${hEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  }

  // ── Paragraph-level run stitching + token replacement ─────────────────────
  function _replaceParagraphTokens(paraXml, tokenMap) {
    const tRegex = /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g;
    const tNodes = [];
    let m;
    while ((m = tRegex.exec(paraXml)) !== null)
      tNodes.push({full:m[0],open:m[1],text:m[2],close:m[3],index:m.index});
    if (!tNodes.length) return paraXml;

    const combined = tNodes.map(n=>n.text).join('');
    let hasToken = Object.keys(tokenMap).some(t => combined.includes(t));
    if (!hasToken) return paraXml;

    let replaced = combined;
    for (const [token, value] of Object.entries(tokenMap))
      while (replaced.includes(token)) replaced = replaced.split(token).join(_esc(value));

    let result = paraXml;
    for (let i=tNodes.length-1; i>=0; i--) {
      const node = tNodes[i];
      const newText = i===0 ? replaced : '';
      let openTag = node.open;
      if (newText && /^ | $/.test(newText) && !openTag.includes('space'))
        openTag = openTag.replace('>', ' xml:space="preserve">');
      result = result.slice(0,node.index) + openTag+newText+node.close + result.slice(node.index+node.full.length);
    }
    return result;
  }

  function _replaceAllTokens(xml, tokenMap) {
    return xml.replace(/(<w:p[ >][\s\S]*?<\/w:p>)/g, p => _replaceParagraphTokens(p, tokenMap));
  }

  // ── Strip <w:sdt> wrappers (depth-aware) ──────────────────────────────────
  function _stripSdtWrappers(xml) {
    const out = []; let i = 0;
    while (i < xml.length) {
      if (xml.slice(i,i+7) === '<w:sdt>') {
        let depth=1, j=i+7;
        while (j<xml.length && depth>0) {
          if      (xml.slice(j,j+7)==='<w:sdt>')  {depth++;j+=7;}
          else if (xml.slice(j,j+8)==='</w:sdt>') {depth--;if(depth>0)j+=8;}
          else j++;
        }
        const block=xml.slice(i,j+8);
        const s=block.indexOf('<w:sdtContent>'), e=block.lastIndexOf('</w:sdtContent>');
        out.push((s!==-1&&e!==-1) ? block.slice(s+14,e) : block);
        i=j+8;
      } else { out.push(xml[i++]); }
    }
    return out.join('');
  }

  // ── Find and expand a repeating template row ──────────────────────────────
  function _expandRows(xml, tokenKey, rows, tokenMapFn, emptyTokens) {
    const pos = xml.indexOf(`{{${tokenKey}}}`);
    if (pos===-1) return xml;
    const before = Math.max(xml.lastIndexOf('<w:tr ',pos), xml.lastIndexOf('<w:tr>',pos));
    const after  = xml.indexOf('</w:tr>',pos);
    if (before===-1||after===-1) return xml;
    const tmplRow = xml.slice(before, after+7);

    let newRows;
    if (rows && rows.length) {
      newRows = rows.map(r => _replaceAllTokens(tmplRow, tokenMapFn(r))).join('');
    } else {
      const em = {}; emptyTokens.forEach(t=>{em[`{{${t}}}`]='';});
      newRows = _replaceAllTokens(tmplRow, em);
    }
    return xml.slice(0,before) + newRows + xml.slice(after+7);
  }

  // ── Main ──────────────────────────────────────────────────────────────────
  async function download(formData, filename) {

    const buf      = await _fetchTemplate();
    const zipFiles = await ZipLib.readZip(buf);
    let   xml      = await ZipLib.getFileText(zipFiles, 'word/document.xml');

    // 1. Strip Word content controls and permission markers
    xml = _stripSdtWrappers(xml);
    xml = xml.replace(/<w:permStart[^>]*\/>/g, '');
    xml = xml.replace(/<w:permEnd[^>]*\/>/g, '');

    // 2. Collect images from RTE overlays
    const imgEntries = [];
    let   rIdCounter = 100;

    function _collectRTEImages(editorId) {
      const images = (typeof RTE !== 'undefined') ? RTE.getImages(editorId) : [];
      let parasXml = '';
      for (const img of images) {
        const rId  = `rId${rIdCounter++}`;
        const ext  = _mimeExt(img.src);
        const wPx  = img.w || 300;
        const hPx  = img.h || 200;
        const ar   = hPx / Math.max(wPx,1);
        const wIn  = Math.min(wPx/96, 6.0);
        const wEmu = Math.round(wIn*914400);
        const hEmu = Math.round(wIn*ar*914400);
        imgEntries.push({ rId, ext, bytes: _dataUrlToBytes(img.src), name: `word/media/img${rId}.${ext}` });
        parasXml += _makeImageXml(wEmu, hEmu, rId);
      }
      return parasXml;
    }

    const summaryImgXml  = _collectRTEImages('rte-summary');
    const followupImgXml = _collectRTEImages('rte-followup');

    // 3. Token map
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

    // 4. Replace tokens
    xml = _replaceAllTokens(xml, tokenMap);

    // 5. Inject image paragraphs after the summary/followup paragraph
    //    Find the </w:tc> that closes the Summary or FollowUp table cell
    function _injectAfterCell(searchToken, imgXml) {
      if (!imgXml) return;
      // The token was replaced — find its table cell by looking for a unique
      // text fragment near where it was, then append images after </w:tc>
      // Fallback: insert before </w:body>
      const textSample = _esc(_rteText('rte-' + (searchToken==='Summary'?'summary':'followup'))).slice(0,15);
      let insertPos = -1;
      if (textSample) {
        const idx = xml.indexOf(textSample);
        if (idx !== -1) {
          const cellEnd = xml.indexOf('</w:tc>', idx);
          if (cellEnd !== -1) insertPos = cellEnd;
        }
      }
      if (insertPos === -1) {
        xml = xml.replace('</w:body>', imgXml + '</w:body>');
      } else {
        xml = xml.slice(0, insertPos) + imgXml + xml.slice(insertPos);
      }
    }

    _injectAfterCell('Summary',  summaryImgXml);
    _injectAfterCell('FollowUp', followupImgXml);

    // 6. Expand labor rows
    xml = _expandRows(xml, 'L.Date', formData.laborRows || [], r => ({
      '{{L.Date}}':       r.date    ||'',
      '{{L.DOW}}':        r.dow     ||'',
      '{{L.Std.Reg}}':    r.std_reg ||'',
      '{{L.Std.OT}}':     r.std_ot  ||'',
      '{{L.Std.Hol}}':    r.std_hol ||'',
      '{{L.Next.Reg}}':   r.nxt_reg ||'',
      '{{L.Next.OT}}':    r.nxt_ot  ||'',
      '{{L.Next.Hol}}':   r.nxt_hol ||'',
      '{{L.Second.Reg}}': r.sec_reg ||'',
      '{{L.Second.OT}}':  r.sec_ot  ||'',
      '{{L.Second.Hol}}': r.sec_hol ||'',
      '{{L.Notes}}':      r.notes   ||'',
    }), ['L.Date','L.DOW','L.Std.Reg','L.Std.OT','L.Std.Hol','L.Next.Reg','L.Next.OT','L.Next.Hol','L.Second.Reg','L.Second.OT','L.Second.Hol','L.Notes']);

    // 7. Expand parts rows
    xml = _expandRows(xml, 'P.Num', formData.partsRows || [], r => ({
      '{{P.Num}}':     r.num         ||'',
      '{{P.Desc}}':    r.desc        ||'',
      '{{P.Serials}}': r.ser         ||'',
      '{{P.Qty}}':     String(r.qty||''),
    }), ['P.Num','P.Desc','P.Serials','P.Qty']);

    // 8. Rebuild zip
    const outFiles = {};
    for (const [name] of Object.entries(zipFiles)) {
      if (name === 'word/document.xml') {
        outFiles[name] = xml;
      } else if (name === 'word/_rels/document.xml.rels' && imgEntries.length) {
        let rels = await ZipLib.getFileText(zipFiles, name);
        for (const img of imgEntries) {
          const rel = `<Relationship Id="${img.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/img${img.rId}.${img.ext}"/>`;
          rels = rels.replace('</Relationships>', rel + '</Relationships>');
        }
        outFiles[name] = rels;
      } else if (name === '[Content_Types].xml' && imgEntries.length) {
        let ct = await ZipLib.getFileText(zipFiles, name);
        for (const ext of [...new Set(imgEntries.map(e=>e.ext))]) {
          if (!ct.includes(`Extension="${ext}"`)) {
            const mime = ext==='jpg'?'image/jpeg':`image/${ext}`;
            ct = ct.replace('</Types>', `<Default Extension="${ext}" ContentType="${mime}"/></Types>`);
          }
        }
        outFiles[name] = ct;
      } else {
        outFiles[name] = await ZipLib.getFileAsBytes(zipFiles, name);
      }
    }
    for (const img of imgEntries) outFiles[img.name] = img.bytes;

    // 9. Write zip and download
    const bytes = await ZipLib.writeZip(outFiles);
    const blob  = new Blob([bytes], {type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
    const url   = URL.createObjectURL(blob);
    const a     = Object.assign(document.createElement('a'), {href:url, download:filename});
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url), 10000);
  }

  return { download };
})();
