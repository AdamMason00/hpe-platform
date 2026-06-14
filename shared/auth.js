/* =====================================================================
 * HPE Service Platform — Authentication
 * ---------------------------------------------------------------------
 * Two tiers:
 *   Tier 1  Google email (managers / admin) — resolved by getUser() from
 *           the Apps Script backend (Session.getActiveUser().getEmail()),
 *           with a client-side roster fallback in HPE_CONFIG.EMAIL_USERS.
 *   Tier 2  PIN (techs / support) — 4-digit PIN matched against the staff
 *           roster loaded from the backend (Staff/PINs).
 *
 * The resolved session is cached in sessionStorage under 'hpe_session' so
 * every module/page shares it without re-authenticating.
 * ===================================================================== */

(function (global) {
  'use strict';

  var CFG = global.HPE_CONFIG || {};
  var API = global.HPE_API;
  var KEY = 'hpe_session';

  /* ---- session storage ------------------------------------------- */
  function getSession() {
    try { return JSON.parse(sessionStorage.getItem(KEY) || 'null'); }
    catch (e) { return null; }
  }
  function setSession(sess) {
    try { sessionStorage.setItem(KEY, JSON.stringify(sess)); } catch (e) {}
    return sess;
  }
  function clearSession() {
    try { sessionStorage.removeItem(KEY); } catch (e) {}
  }

  /* ---- role helpers ---------------------------------------------- */
  function isAdmin(sess)   { return !!sess && sess.role === 'admin'; }
  function isManager(sess) { return !!sess && (sess.role === 'manager' || sess.role === 'admin'); }
  function isTech(sess)    { return !!sess && sess.role === 'tech'; }
  function isSupport(sess) { return !!sess && sess.role === 'support'; }

  /* ---- Tier 1: Google email login -------------------------------
   * Asks the backend who the active Google user is. If the backend is
   * unreachable, falls back to a manually-entered email checked against
   * the local roster so managers are never locked out.
   * --------------------------------------------------------------- */
  function loginWithGoogle() {
    if (!API) return Promise.reject(new Error('API not loaded'));
    return API.getUser().then(function (res) {
      var email = (res && (res.email || (res.user && res.user.email))) || '';
      return resolveEmail(email);
    });
  }

  // Resolve an email -> session (used by Google login and the fallback).
  function resolveEmail(email) {
    email = (email || '').toString().trim().toLowerCase();
    if (!email) throw new Error('No Google account detected. Use the manual sign-in or a PIN.');
    var u = CFG.userForEmail(email);
    if (!u) throw new Error('“' + email + '” is not authorized as a manager. Ask an admin to add you, or use a PIN.');
    return setSession({
      tier: 'email',
      email: email,
      name: u.name,
      role: u.role,            // 'admin' | 'manager'
      store: u.store,          // 'south' | 'north' | null
      ts: Date.now()
    });
  }

  /* ---- Tier 2: PIN login ----------------------------------------
   * Loads the staff roster (with PINs) from the backend and matches the
   * entered 4-digit PIN. Falls back to HPE_CONFIG.DEFAULT_STAFF only if
   * those defaults have PINs assigned (normally they won't until an admin
   * sets them, so a backend round-trip is expected).
   * --------------------------------------------------------------- */
  function loginWithPIN(pin) {
    pin = (pin || '').toString().trim();
    if (!/^\d{4}$/.test(pin)) {
      return Promise.reject(new Error('Enter a 4-digit PIN.'));
    }
    var loadStaff = API ? API.loadStaff().then(function (r) {
      return (r && (r.staff || r.data)) || [];
    }).catch(function () { return CFG.DEFAULT_STAFF || []; })
      : Promise.resolve(CFG.DEFAULT_STAFF || []);

    return loadStaff.then(function (staff) {
      var match = null;
      for (var i = 0; i < staff.length; i++) {
        if (staff[i] && String(staff[i].pin || '').trim() === pin) { match = staff[i]; break; }
      }
      if (!match) throw new Error('PIN not recognized. Check with your manager.');
      var queues = (match.queue ? [match.queue] : (CFG.queuesForStaff ? CFG.queuesForStaff(match.name) : []));
      return setSession({
        tier: 'pin',
        name: match.name,
        role: match.roleType || 'tech',   // 'tech' | 'support' | 'admin'
        store: match.store,
        division: match.division,
        fte: match.fte,
        queues: queues,
        ts: Date.now()
      });
    });
  }

  /* ---- guard: require a session on protected pages --------------- */
  function requireSession(redirectTo) {
    var sess = getSession();
    if (!sess) {
      var base = redirectTo || '../index.html';
      window.location.href = base;
      return null;
    }
    return sess;
  }

  function logout(redirectTo) {
    clearSession();
    if (redirectTo) window.location.href = redirectTo;
  }

  global.HPE_AUTH = {
    getSession: getSession,
    setSession: setSession,
    clearSession: clearSession,
    loginWithGoogle: loginWithGoogle,
    resolveEmail: resolveEmail,
    loginWithPIN: loginWithPIN,
    requireSession: requireSession,
    logout: logout,
    isAdmin: isAdmin,
    isManager: isManager,
    isTech: isTech,
    isSupport: isSupport
  };

})(typeof window !== 'undefined' ? window : this);
