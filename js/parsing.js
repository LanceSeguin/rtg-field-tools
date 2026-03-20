// ─────────────────────────────────────────────────────────────────────────────
// parsing.js — Parse Outlook event subject + body into work order fields
//
// This mirrors the logic in parsing.py from the original desktop app.
// Edit this file to fix or improve how fields are extracted from calendar events.
//
// INPUT:  subject (string), body (string — may be HTML from Graph API)
// OUTPUT: object with keys matching the work order form fields
// ─────────────────────────────────────────────────────────────────────────────

const Parsing = (() => {

  // ── HTML body stripping ───────────────────────────────────────────────────
  // Graph API returns body as HTML when bodyType is 'html'. Strip tags for parsing.
  function stripHtml(html) {
    if (!html) return '';
    return html
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n')
      .replace(/<\/div>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .trim();
  }

  // ── Normalization ─────────────────────────────────────────────────────────
  function normalizeWs(text) {
    return (text || '')
      .replace(/\u00A0|\u2007|\u202F/g, ' ')  // non-breaking spaces
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/\t/g, ' ');
  }

  // ── Section header detection ──────────────────────────────────────────────
  const SECTION_HEADERS = [
    'Site Contact', 'Order Contact', 'Location Contact',
    'Site Address', 'Scope', 'Purchase Order', 'PO',
    'Site', 'Customer', 'Service Technician', 'Service Agency',
    'Service Agency Order', 'Product Line', 'System/Serial Number',
  ];

  function isHeaderLine(line) {
    const l = (line || '').trim();
    return SECTION_HEADERS.some(h =>
      new RegExp(`^${h.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*[:\\-\\u2013\\u2014]?\\s*`, 'i').test(l)
    );
  }

  // ── Block extractor — lines under a section header ────────────────────────
  function getBlockLines(body, header) {
    const rx = new RegExp(`^\\s*${header.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*[:\\-]?\\s*$`, 'im');
    const m  = rx.exec(body);
    if (!m) return [];

    const lines  = body.slice(m.index + m[0].length).split('\n');
    const out    = [];
    let blanks   = 0;

    for (const line of lines) {
      if (!line.trim()) {
        blanks++;
        if (blanks >= 2) break;
        continue;
      }
      if (isHeaderLine(line)) break;
      blanks = 0;
      out.push(line.trimEnd());
      if (out.length >= 60) break;
    }
    return out;
  }

  // ── Contact parsing ───────────────────────────────────────────────────────
  const PHONE_RE = /(?:\+?\d[\d\s().\-]{6,}\d(?:\s*(?:x|ext\.?)\s*\d{1,6})?)/;
  const EMAIL_RE = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/;
  const LABEL_RE = /^(?:name|phone|email|number|contact(?:\s*name)?)\s*[:\-]\s*/i;

  function looksLikeName(text) {
    const t = (text || '').trim();
    if (!t || t.includes('@')) return false;
    if (t.length < 2) return false;
    return /^[A-Za-z .'\-\u2019]+$/.test(t);
  }

  function parseContactFromLines(lines) {
    const out = { ContactName: '', ContactPhone: '', ContactEmail: '' };
    if (!lines.length) return out;

    const block = lines.join('\n');
    const em = EMAIL_RE.exec(block);
    if (em) out.ContactEmail = em[0].trim();

    const ph = PHONE_RE.exec(block);
    if (ph) out.ContactPhone = ph[0].trim();

    for (const line of lines) {
      const c = line.trim().replace(LABEL_RE, '');
      if (!c) continue;
      if (out.ContactEmail && c === out.ContactEmail) continue;
      if (out.ContactPhone && c === out.ContactPhone) continue;
      if (looksLikeName(c)) { out.ContactName = c; break; }
    }
    return out;
  }

  function parseContact(body) {
    const site  = parseContactFromLines(getBlockLines(body, 'Site Contact'));
    const order = parseContactFromLines(getBlockLines(body, 'Order Contact'));
    const loc   = parseContactFromLines(getBlockLines(body, 'Location Contact'));
    // Merge: site first, fill gaps from order then location
    const merge = (dst, src) => {
      Object.keys(src).forEach(k => { if (!dst[k] && src[k]) dst[k] = src[k]; });
      return dst;
    };
    return merge(merge(site, order), loc);
  }

  // ── Address normalizer ────────────────────────────────────────────────────
  // Tries to produce exactly 2 lines: street \n City, ST ZIP
  function normalizeAddress(raw) {
    if (!raw) return '';
    const lines = raw.split('\n').map(l => l.trim()).filter(Boolean);
    if (lines.length >= 2) return `${lines[0]}\n${lines[1]}`;

    const one = lines[0] || raw.trim();
    const m   = one.match(/^(.*?)[,\s]+([A-Za-z .'\-]+)[,\s]+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/);
    if (m && m[1] && m[2]) {
      return `${m[1].replace(/,\s*$/, '').trim()}\n${m[2].trim()}, ${m[3]} ${m[4]}`;
    }
    return raw;
  }

  // ── Section capture (same-line or under-header) ───────────────────────────
  function captureSection(body, header, maxLines = 60) {
    // Try same-line: "Header: value"
    const rxSame = new RegExp(`^\\s*${header.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*[:\\-]\\s*(.+)$`, 'im');
    const mSame  = rxSame.exec(body);

    // Try alone: "Header" on its own line
    const rxAlone = new RegExp(`^\\s*${header.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*[:\\-]?\\s*$`, 'im');
    const mAlone  = rxAlone.exec(body);

    if (!mSame && !mAlone) return '';

    // Use whichever appears first in the text
    let kind, startIdx, firstLine = '';
    if (mSame && mAlone) {
      if (mSame.index <= mAlone.index) { kind='same'; startIdx=mSame.index+mSame[0].length; firstLine=mSame[1].trim(); }
      else { kind='alone'; startIdx=mAlone.index+mAlone[0].length; }
    } else if (mSame)  { kind='same';  startIdx=mSame.index+mSame[0].length;  firstLine=mSame[1].trim(); }
    else               { kind='alone'; startIdx=mAlone.index+mAlone[0].length; }

    const outLines = firstLine ? [firstLine] : [];
    let blanks = 0;

    for (const line of body.slice(startIdx).split('\n')) {
      if (!line.trim()) {
        blanks++;
        if (blanks >= 2) break;
        continue;
      }
      if (isHeaderLine(line)) break;
      blanks = 0;
      outLines.push(line.trimEnd());
      if (outLines.length >= maxLines) break;
    }

    return outLines.join('\n').trim();
  }

  // ── PO parser ─────────────────────────────────────────────────────────────
  const HEADER_WORDS = ['site','site address','site contact','order contact','scope','location','address'];

  function looksLikePO(val) {
    const s = (val || '').trim();
    if (!s || s.length < 2 || s.length > 64) return false;
    if (!/\d/.test(s)) return false;
    if (HEADER_WORDS.some(w => new RegExp(`^${w}\\b`, 'i').test(s))) return false;
    if (/^(tbd|n\/?a|none|na)$/i.test(s)) return false;
    return /^[A-Za-z0-9][A-Za-z0-9 _\-./#]{0,63}$/.test(s);
  }

  function parsePO(body) {
    const clean = v => (v || '').trim().replace(/^[ <>[\](){}.,;]+|[ <>[\](){}.,;]+$/g, '');
    const alias = '(?:PO|P\\.?O\\.?|Purchase\\s*Order(?:\\s*(?:Number|#))?)';
    const sep   = '[:\\-\\u2013\\u2014\\uFF1A]';

    let m = new RegExp(`(?:^|\\n)\\s*${alias}\\s*${sep}\\s*(?<val>\\S.+)`, 'i').exec(body);
    if (m) { const v = clean(m.groups.val); if (looksLikePO(v)) return v; }

    m = /\bPO\s*#\s*([A-Za-z0-9\-_/.]{2,})\b/i.exec(body);
    if (m) { const v = clean(m[1]); if (looksLikePO(v)) return v; }

    return '';
  }

  // ── Main entry point ──────────────────────────────────────────────────────
  /**
   * Parse an Outlook event subject + body into work order fields.
   * @param {string} subject
   * @param {string} body  — may be HTML (will be stripped) or plain text
   * @returns {object}     — keys match form field IDs
   */
  function parseEventFields(subject, body) {
    const out  = {};
    const subj = (subject || '').trim();
    const bod  = normalizeWs(stripHtml(body || ''));

    // ── Subject: "Customer - Technician - SLORDxxxxxx" ────────────────────
    const parts = subj.split(' - ').map(s => s.trim());
    if (parts[0]) out.CustomerName       = parts[0];
    if (parts[1]) out.ServiceTechnician  = parts[1];

    const slord = subj.match(/\bSLORD\d{6}\b/i);
    if (slord) out.ServiceAgencyOrder = slord[0].toUpperCase();
    else if (parts[2] && /^SLORD\d{6}$/i.test(parts[2]))
      out.ServiceAgencyOrder = parts[2].toUpperCase();

    // Site name from subject
    const site = subj.match(/Site\s*[:\-]\s*([^\n;\-[\]()]{1,100})/i);
    if (site) out.ServiceAgency = site[1].trim().replace(/[;\-\n]+$/, '');

    // ── Body: contacts ────────────────────────────────────────────────────
    const contact = parseContact(bod);
    Object.assign(out, contact);

    // ── Body: address & scope ─────────────────────────────────────────────
    const rawAddr = captureSection(bod, 'Site Address', 40);
    if (rawAddr) out.ServiceAddress = normalizeAddress(rawAddr);

    const rawScope = captureSection(bod, 'Scope', 120);
    if (rawScope) out.Scope = rawScope;

    // ── Body: PO ──────────────────────────────────────────────────────────
    out.PONumber = parsePO(bod);

    return out;
  }

  return { parseEventFields, stripHtml };
})();
