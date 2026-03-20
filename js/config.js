// ─────────────────────────────────────────────────────────────────────────────
// config.js — Edit this file to change app settings
// ─────────────────────────────────────────────────────────────────────────────

const CONFIG = {
  // ── Azure App Registration ──────────────────────────────────────────────────
  // Found in portal.azure.com → App registrations → WorkOrder Graph Integration
  clientId:    '8590561e-9ae0-4988-9889-c994688f8db2',
  tenantId:    'eb06985d-06ca-4a17-81da-629ab99f6505',
  redirectUri: 'https://lanceseguin.github.io/rtg-field-tools/',

  // ── Microsoft Graph scopes ──────────────────────────────────────────────────
  // Add 'Calendars.Read.Shared' here once IT grants that permission
  scopes: [
    'Calendars.Read',
    'User.Read',
    // 'Calendars.Read.Shared',  // ← uncomment when IT grants this
  ],

  // ── App defaults ────────────────────────────────────────────────────────────
  defaultTechnician:  'Seguin',          // pre-fills Service Technician field
  defaultAgency:      'Customer Name',   // pre-fills Service Agency field
  calendarLookbackDays:  7,              // how many days back to default the calendar start
  calendarLookaheadDays: 30,             // how many days ahead to default the calendar end
};
