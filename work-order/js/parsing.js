// ─────────────────────────────────────────────────────────────────────────────
// parsing.js — Parse Outlook event subject + body into work order fields
//
// Written to match the exact RTG invite format:
//
//   Subject:  Stolthaven - Lance - SLORD460215
//
//   Site Contact:
//       Jimmy D'Anna
//       Stolthaven
//       504-232-4031
//       j.d.anna@stolt.com
//
//   Order Contact:
//       Brad Plattsmier
//       John H Carter
//       504-232-1876
//       brad.plattsmier@johnhcarter.com
//
//   Purchase Order:
//       JHC1181753
//
//   Location:
//       2444 English Turn Road
//       Braithwaite, LA 70040
//
//   Scope of Work:
//       Upgrade existing TM server to 6G4...
//
// Edit this file to fix or improve field extraction.
// ─────────────────────────────────────────────────────────────────────────────

const Parsing = (() => {

  // ── Strip HTML tags from Graph API body ──────────────────────────────────
  function stripHtml(html) {
    if (!html) return '';
    return html
      .replace(/<br\s*\/?>/gi,  '\n')
      .replace(/<\/p>/gi,       '\n')
      .replace(/<\/div>/gi,     '\n')
      .replace(/<\/li>/gi,      '\n')
      .replace(/<[^>]+>/g,      '')
      .replace(/&nbsp;/g,       ' ')
      .replace(/&amp;/g,        '&')
      .replace(/&lt;/g,         '<')
      .replace(/&gt;/g,         '>')
      .replace(/&quot;/g,       '"')
      .replace(/&#39;/g,        "'")
      .replace(/&#x27;/g,       "'")
      .replace(/\u00A0/g,       ' ')   // non-breaking space
      .replace(/\r\n/g,         '\n')
      .replace(/\r/g,           '\n')
      .trim();
  }

  // ── Get all lines under a section header ─────────────────────────────────
  // Matches headers like "Site Contact:", "Scope of Work:", "Location:", etc.
  // Stops at the next header or two consecutive blank lines.
  function getSection(text, ...headerNames) {
    for (const header of headerNames) {
      // Match "Header:" or "Header :" at the start of a line (case-insensitive)
      const rx = new RegExp(
        `(?:^|\\n)[ \\t]*${header.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}[ \\t]*:[ \\t]*\\n?`,
        'i'
      );
      const m = rx.exec(text);
      if (!m) continue;

      const after = text.slice(m.index + m[0].length);
      const lines = after.split('\n');
      const out   = [];
      let blanks  = 0;

      for (const line of lines) {
        const trimmed = line.trim();
        // Stop at next section header (word followed by colon at line start)
        if (/^[A-Z][A-Za-z ]+:/.test(trimmed) && trimmed.length < 60) break;
        if (!trimmed) {
          blanks++;
          if (blanks >= 2) break;
          continue;
        }
        blanks = 0;
        out.push(trimmed);
        if (out.length >= 80) break;
      }

      if (out.length) return out;
    }
    return [];
  }

  // ── Patterns ──────────────────────────────────────────────────────────────
  const EMAIL_RE = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/;
  const PHONE_RE = /(?:\+?1[\s.\-]?)?\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4}(?:\s*(?:x|ext\.?)\s*\d{1,6})?/;

  // ── Parse contact block lines into name/phone/email ───────────────────────
  function parseContact(lines) {
    const out = { name: '', phone: '', email: '' };
    if (!lines.length) return out;

    const block = lines.join('\n');

    const em = EMAIL_RE.exec(block);
    if (em) out.email = em[0].trim();

    const ph = PHONE_RE.exec(block);
    if (ph) out.phone = ph[0].trim();

    // Name = first line that isn't a phone or email and looks like a person's name
    for (const line of lines) {
      const l = line.trim();
      if (!l) continue;
      if (out.email && l.includes('@')) continue;
      if (out.phone && PHONE_RE.test(l)) continue;
      // Accept if it looks like a name (letters, spaces, apostrophes, hyphens)
      if (/^[A-Za-z][A-Za-z '.'\-]{1,40}$/.test(l)) {
        out.name = l;
        break;
      }
    }

    return out;
  }

  // ── Main entry point ──────────────────────────────────────────────────────
  /**
   * Parse an Outlook event subject + body into work order fields.
   *
   * @param {string} subject  — event subject line
   * @param {string} body     — event body (may be HTML from Graph API)
   * @returns {object}        — keys matching form field IDs
   */
  function parseEventFields(subject, body) {
    const out  = {};
    const subj = (subject || '').trim();
    const text = stripHtml(body || '');

    // ── Subject: "Customer - Technician - SLORDxxxxxx" ───────────────────
    const parts = subj.split(' - ').map(s => s.trim()).filter(Boolean);
    if (parts[0]) out.CustomerName      = parts[0];
    if (parts[1]) out.ServiceTechnician = parts[1];

    const slord = subj.match(/\bSLORD\d+\b/i);
    if (slord) out.ServiceAgencyOrder = slord[0].toUpperCase();
    else if (parts[2] && /^SLORD\d+$/i.test(parts[2]))
      out.ServiceAgencyOrder = parts[2].toUpperCase();

    // ── Site Contact ──────────────────────────────────────────────────────
    const siteLines = getSection(text, 'Site Contact');
    if (siteLines.length) {
      const c = parseContact(siteLines);
      if (c.name)  out.ContactName  = c.name;
      if (c.phone) out.ContactPhone = c.phone;
      if (c.email) out.ContactEmail = c.email;
    }

    // ── Order Contact (fill any gaps left by Site Contact) ────────────────
    const orderLines = getSection(text, 'Order Contact');
    if (orderLines.length) {
      const c = parseContact(orderLines);
      if (!out.ContactName  && c.name)  out.ContactName  = c.name;
      if (!out.ContactPhone && c.phone) out.ContactPhone = c.phone;
      if (!out.ContactEmail && c.email) out.ContactEmail = c.email;
    }

    // ── Purchase Order ────────────────────────────────────────────────────
    const poLines = getSection(text, 'Purchase Order', 'PO');
    if (poLines.length) {
      const val = poLines[0].trim();
      // Must contain a digit and not be a section label
      if (/\d/.test(val) && val.length <= 64) out.PONumber = val;
    }

    // ── Location → Service Address ────────────────────────────────────────
    // Your invites use "Location:" for the site address
    const locLines = getSection(text, 'Location', 'Site Address', 'Site');
    if (locLines.length) {
      // Normalize to max 2 lines: street \n city, state zip
      out.ServiceAddress = locLines.slice(0, 2).join('\n');
    }

    // ── Scope of Work → Scope ─────────────────────────────────────────────
    // Your invites use "Scope of Work:" 
    const scopeLines = getSection(text, 'Scope of Work', 'Scope');
    if (scopeLines.length) {
      out.Scope = scopeLines.join('\n');
    }

    // ── Service Agency from subject Site field ────────────────────────────
    const sitePart = subj.match(/Site\s*[:\-]\s*([^\n\-[\]()]{1,80})/i);
    if (sitePart) out.ServiceAgency = sitePart[1].trim();

    return out;
  }

  return { parseEventFields, stripHtml };
})();
