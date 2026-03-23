// ─────────────────────────────────────────────────────────────────────────────
// config.js — Shared configuration for all RTG Field Tools
// Edit this file to change Azure credentials or app defaults.
// ─────────────────────────────────────────────────────────────────────────────

const CONFIG = {
  // ── Azure App Registration ─────────────────────────────────────────────────
  clientId:    '8590561e-9ae0-4988-9889-c994688f8db2',
  tenantId:    'eb06985d-06ca-4a17-81da-629ab99f6505',

  // Dynamically use the current page origin + path as redirect URI
  // This handles: root hub, /work-order/, /expense/ all with one registration
  // You must add ALL of these to Azure portal → Authentication → Redirect URIs:
  //   https://lanceseguin.github.io/rtg-field-tools/
  //   https://lanceseguin.github.io/rtg-field-tools/work-order/
  //   https://lanceseguin.github.io/rtg-field-tools/expense/
  get redirectUri() {
    return window.location.origin + window.location.pathname;
  },

  // ── Microsoft Graph scopes ─────────────────────────────────────────────────
  scopes: [
    'Calendars.Read',
    'User.Read',
    // 'Calendars.Read.Shared',  // uncomment when IT grants this
  ],

  // ── Work Order defaults ────────────────────────────────────────────────────
  defaultTechnician:     'Seguin',
  defaultAgency:         'Customer Name',
  calendarLookbackDays:  7,
  calendarLookaheadDays: 30,
};
