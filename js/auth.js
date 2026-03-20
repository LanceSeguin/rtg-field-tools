// ─────────────────────────────────────────────────────────────────────────────
// auth.js — Microsoft OAuth 2.0 PKCE login flow
// You should rarely need to edit this file.
// ─────────────────────────────────────────────────────────────────────────────

const Auth = (() => {
  let _token = null;
  let _user  = null;

  // ── PKCE helpers ─────────────────────────────────────────────────────────────
  function _b64url(buf) {
    return btoa(String.fromCharCode(...new Uint8Array(buf)))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  }

  async function _genPKCE() {
    const verifier  = _b64url(crypto.getRandomValues(new Uint8Array(32)));
    const digest    = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
    const challenge = _b64url(digest);
    return { verifier, challenge };
  }

  // ── Public API ────────────────────────────────────────────────────────────────

  /** Redirect user to Microsoft login page */
  async function login() {
    const { verifier, challenge } = await _genPKCE();
    const state = _b64url(crypto.getRandomValues(new Uint8Array(16)));

    sessionStorage.setItem('pkce_verifier', verifier);
    sessionStorage.setItem('pkce_state',    state);

    const params = new URLSearchParams({
      client_id:             CONFIG.clientId,
      response_type:         'code',
      redirect_uri:          CONFIG.redirectUri,
      scope:                 CONFIG.scopes.join(' ') + ' offline_access',
      state,
      code_challenge:        challenge,
      code_challenge_method: 'S256',
      response_mode:         'fragment',
    });

    location.href = `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/authorize?${params}`;
  }

  /** Call this on page load — handles the redirect back from Microsoft */
  async function handleRedirect() {
    const hash = location.hash.substring(1);
    if (!hash) return false;

    const p    = new URLSearchParams(hash);
    const code = p.get('code');
    if (!code) return false;

    const verifier = sessionStorage.getItem('pkce_verifier');
    if (!verifier) return false;

    try {
      const resp = await fetch(
        `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`,
        {
          method:  'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body:    new URLSearchParams({
            client_id:     CONFIG.clientId,
            grant_type:    'authorization_code',
            code,
            redirect_uri:  CONFIG.redirectUri,
            code_verifier: verifier,
          }),
        }
      );

      const data = await resp.json();
      if (data.access_token) {
        _token = data.access_token;
        sessionStorage.setItem('rtg_token',   _token);
        if (data.refresh_token)
          sessionStorage.setItem('rtg_refresh', data.refresh_token);
        history.replaceState({}, '', location.pathname);
        return true;
      }
    } catch (e) {
      console.error('Token exchange failed', e);
    }
    return false;
  }

  /** Restore token from sessionStorage (survives page refresh within tab) */
  function restoreSession() {
    _token = sessionStorage.getItem('rtg_token');
    return !!_token;
  }

  /** Get the current access token */
  function getToken() {
    return _token || sessionStorage.getItem('rtg_token');
  }

  /** Fetch the signed-in user's profile from Graph */
  async function getUser() {
    if (_user) return _user;
    const token = getToken();
    if (!token) return null;
    try {
      const r = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (r.ok) { _user = await r.json(); return _user; }
    } catch {}
    return null;
  }

  /** Sign out — clears session storage and resets state */
  function logout() {
    sessionStorage.clear();
    _token = null;
    _user  = null;
  }

  return { login, handleRedirect, restoreSession, getToken, getUser, logout };
})();
