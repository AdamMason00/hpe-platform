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

  /* ---- roster loader (backend, falling back to local defaults) --- */
  function loadRoster() {
    if (!API) return Promise.resolve(CFG.DEFAULT_STAFF || []);
    return API.loadStaff().then(function (r) {
      var s = (r && (r.staff || r.data)) || [];
      return s.length ? s : (CFG.DEFAULT_STAFF || []);
    }).catch(function () { return CFG.DEFAULT_STAFF || []; });
  }

  function findByEmail(staff, email) {
    email = email.toLowerCase();
    for (var i = 0; i < staff.length; i++) {
      if (staff[i] && String(staff[i].email || '').toLowerCase() === email) return staff[i];
    }
    return null;
  }
  function findByName(staff, name) {
    name = String(name || '').toLowerCase();
    for (var i = 0; i < staff.length; i++) {
      if (staff[i] && String(staff[i].name || '').toLowerCase() === name) return staff[i];
    }
    return null;
  }
  function findByEmpNum(staff, num) {
    num = String(num || '').trim().toUpperCase();
    for (var i = 0; i < staff.length; i++) {
      if (staff[i] && String(staff[i].empNum || '').trim().toUpperCase() === num && num) return staff[i];
    }
    return null;
  }

  /* ---- Unified login --------------------------------------------
   * identifier = HPE email (managers/admin) OR employee number (techs/
   * support). passcode = the person's PIN. The passcode is only enforced
   * when the person actually has a PIN set, so nobody is locked out before
   * PINs are assigned. Returns the new session.
   * --------------------------------------------------------------- */
  function login(identifier, passcode) {
    identifier = (identifier || '').toString().trim();
    passcode   = (passcode   || '').toString().trim();
    if (!identifier) return Promise.reject(new Error('Enter your HPE email or employee number.'));
    return loadRoster().then(function (staff) {
      return (identifier.indexOf('@') !== -1)
        ? emailLogin(identifier.toLowerCase(), passcode, staff)
        : empNumLogin(identifier, passcode, staff);
    });
  }

  function checkPasscode(rec, passcode) {
    var pin = rec ? String(rec.pin || '').trim() : '';
    if (pin && passcode !== pin) throw new Error('Incorrect passcode.');
  }

  function emailLogin(email, passcode, staff) {
    var u = CFG.userForEmail(email);                 // privileged role/store, if any
    var rec = findByEmail(staff, email) || (u && findByName(staff, u.name)) || null;
    if (!u && !rec) throw new Error('“' + email + '” is not authorized. Check the address, or sign in with your employee number.');
    checkPasscode(rec, passcode);
    return setSession({
      tier: 'email', email: email,
      name:  u ? u.name  : rec.name,
      role:  u ? u.role  : (rec.roleType || 'manager'),
      store: u ? u.store : rec.store,
      empNum: rec ? (rec.empNum || '') : '',
      ts: Date.now()
    });
  }

  function empNumLogin(num, passcode, staff) {
    var rec = findByEmpNum(staff, num);
    if (!rec) throw new Error('Employee number “' + num + '” not recognized. Check with your manager.');
    checkPasscode(rec, passcode);
    var queues = (rec.queue ? [rec.queue] : (CFG.queuesForStaff ? CFG.queuesForStaff(rec.name) : []));
    return setSession({
      tier: 'empnum', name: rec.name,
      role: rec.roleType || 'tech', store: rec.store, division: rec.division,
      fte: rec.fte, queues: queues, empNum: rec.empNum, ts: Date.now()
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
    login: login,
    requireSession: requireSession,
    logout: logout,
    isAdmin: isAdmin,
    isManager: isManager,
    isTech: isTech,
    isSupport: isSupport
  };

})(typeof window !== 'undefined' ? window : this);
