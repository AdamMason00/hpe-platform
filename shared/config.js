/* =====================================================================
 * HPE Service Platform — Shared Configuration
 * Hyde Park Equipment (Kubota dealership) — South & North stores
 * ---------------------------------------------------------------------
 * Single source of truth for: backend wiring, auth roster, store list,
 * KPI thresholds, default employee roster, and support WO assignments.
 * Loaded by every module (login hub, KPI manager, warranty).
 * ===================================================================== */

(function (global) {
  'use strict';

  /* ---------------------------------------------------------------
   * BACKEND WIRING
   * The Apps Script Web App URL is shared by all modules so every
   * page talks to the same deployment. Sheet / Drive IDs are kept
   * here for reference and are also defined server-side in Code.gs.
   * ------------------------------------------------------------- */
  var BACKEND = {
    API_URL: 'https://script.google.com/macros/s/AKfycbwkniPdgD25XWiqYaJR7UfXZgp8vmfQ6s3xjRMtuHMEsLQKrWhW-yf29_7WUmzgKBtkYQ/exec',
    SHEET_ID: '1Ljh-Ycf1ut6TyV2NgXFRrypRUJzLRPpdUjl_yOIHrBw',
    DRIVE_FOLDER: '1m9wv8eaWhAaLe1qZ0T0P35zUmt4NrSPQ',   // uploads + backups
    PHOTOS_FOLDER: '1kbsKqfQp-Ms4YqwOtWxD2p3JYTiPEUu5'    // claim photos
  };

  /* ---------------------------------------------------------------
   * STORES
   * Labour/POS "Division" codes map to stores: S = South, M = North.
   * ------------------------------------------------------------- */
  var STORES = [
    { id: 'south', name: 'South Store', division: 'S', color: '#E8620A' },
    { id: 'north', name: 'North Store', division: 'M', color: '#2563EB' }
  ];

  function storeByDivision(div) {
    div = (div || '').toString().trim().toUpperCase();
    for (var i = 0; i < STORES.length; i++) {
      if (STORES[i].division === div) return STORES[i];
    }
    return null;
  }

  /* ---------------------------------------------------------------
   * AUTH — Tier 1: Google email (managers / admin)
   * Roles: admin (full), manager (own store).
   * ------------------------------------------------------------- */
  var EMAIL_USERS = {
    'adam@hydeparkequipment.ca':         { name: 'Adam',          role: 'admin',   store: null },
    'bapfelbeck@hydeparkequipment.ca':   { name: 'B. Apfelbeck',  role: 'admin',   store: null },
    'johnwilliams@hydeparkequipment.ca': { name: 'John Williams', role: 'admin',   store: null },
    'steve@hydeparkequipment.ca':        { name: 'Steve Hayes',   role: 'manager', store: 'south' },
    'bill@hydeparkequipment.ca':         { name: 'Bill Denison',  role: 'manager', store: 'north' }
  };

  function userForEmail(email) {
    if (!email) return null;
    return EMAIL_USERS[email.toString().trim().toLowerCase()] || null;
  }

  /* ---------------------------------------------------------------
   * KPI CONFIGURATION DEFAULTS
   * Admin-editable on the Configuration page; persisted to backend.
   * ------------------------------------------------------------- */
  var KPI_CONFIG = {
    techEff: 75,        // % — KPI1 efficiency target
    comeback: 2,        // % — KPI1 comeback ceiling
    svcGm: 78,          // % — KPI2 service gross margin floor
    partsGm: 32,        // % — KPI3 parts gross margin floor
    wipMax: 2,          // % of revenue — KPI4 open WIP ceiling
    cap: 9000,          // $ — quarterly tech bonus pool cap (live value)
    techShare: 73,      // % — tech portion of the pool (support = 27%)
    growthRate: 30,     // % — growth bank accrual rate
    warrantyAnnual: 5000,  // $ — warranty annual bonus pool (legacy)
    topUpPool: 50000,      // $ — year-end top-up pool (legacy)
    // WO ageing thresholds (days) for the open-WO flagging
    woWarnDays: 30,
    woCriticalDays: 60,
    // Efficiency WO flag bounds
    effHighFlag: 100,   // > 100% flagged (possible over-billing)
    effLowFlag: 75      // < 75% flagged (under target)
  };

  // Manager annual KPI-bonus caps (year-end growth handled separately).
  var MANAGER_BONUS = {
    'steve@hydeparkequipment.ca': { name: 'Steve Hayes',  store: 'south', annualCap: 6000 },
    'bill@hydeparkequipment.ca':  { name: 'Bill Denison', store: 'north', annualCap: 6000 }
  };

  /* ---------------------------------------------------------------
   * DEFAULT EMPLOYEE ROSTER  (fallback only)
   * Real roster, names, stores, roles, FTE and pay reflect the live
   * HydePark_KPI_Database_2026 sheet (legacy AppData blob, v5.0).
   * PINs are intentionally BLANK here — real PINs live only in the
   * private backend Staff blob, never in this public repo. They are
   * carried over by Code.gs migrateFromLegacy(). The frontend normally
   * loads the roster from the backend; this array is just the offline
   * fallback. roleType drives dashboards:
   *   'tech'    -> efficiency-based KPI dashboard
   *   'support' -> open-WO queue dashboard
   *   'admin'   -> warranty admin / office admin
   *   'manager' -> store dashboard (also email-auth Tier 1)
   * division: S = South, M = North.  active:false = no longer employed.
   * ------------------------------------------------------------- */
  var DEFAULT_STAFF = [
    // ---- South techs ----
    { name: 'Jared Lobban',     store: 'south', division: 'S', roleType: 'tech',    fte: 1.0, pin: '', empNum: 'S043', payRate: 43,    payType: 'Hourly', active: true },
    { name: 'John Reiger',      store: 'south', division: 'S', roleType: 'tech',    fte: 1.0, pin: '', empNum: 'S077', payRate: 36,    payType: 'Hourly', active: true },
    { name: 'Tyler Nicholson',  store: 'south', division: 'S', roleType: 'tech',    fte: 1.0, pin: '', empNum: 'S079', payRate: 32,    payType: 'Hourly', active: true },
    { name: 'Caden Koetsier',   store: 'south', division: 'S', roleType: 'tech',    fte: 1.0, pin: '', empNum: 'S076', payRate: 26,    payType: 'Hourly', active: true },
    { name: 'Ed Kindt',         store: 'south', division: 'S', roleType: 'tech',    fte: 1.0, pin: '', empNum: 'S036', payRate: 35.75, payType: 'Hourly', active: true },
    { name: 'Al Monck',         store: 'south', division: 'S', roleType: 'tech',    fte: 1.0, pin: '', empNum: 'S009', payRate: 28.5,  payType: 'Hourly', active: true },
    { name: 'Logan Hardman',    store: 'south', division: 'S', roleType: 'tech',    fte: 0.5, pin: '', empNum: 'S080', payRate: 19,    payType: 'Hourly', active: true },
    { name: 'Tyler Mcdougall',  store: 'south', division: 'S', roleType: 'tech',    fte: 0.5, pin: '', empNum: '',     payRate: 19,    payType: 'Hourly', active: true },
    // ---- South support ----
    { name: 'Stew Brooks',      store: 'south', division: 'S', roleType: 'support', fte: 1.0, pin: '', payRate: 27,    payType: 'Hourly', active: true, queue: 'WS' },
    { name: 'Amy',              store: 'south', division: 'S', roleType: 'support', fte: 1.0, pin: '', payRate: 30.5,  payType: 'Hourly', active: true, queue: 'IS' },
    { name: 'Matt Berkelmans',  store: 'south', division: 'S', roleType: 'support', fte: 1.0, pin: '', payRate: 28.5,  payType: 'Hourly', active: true, queue: 'IS' },
    { name: 'Thorin Wilson',    store: 'south', division: 'S', roleType: 'support', fte: 1.0, pin: '', payRate: 29,    payType: 'Hourly', active: true, queue: 'IS' },
    // ---- North techs ----
    { name: 'Alex Beer',        store: 'north', division: 'M', roleType: 'tech',    fte: 1.0, pin: '', empNum: 'S070', payRate: 40,    payType: 'Hourly', active: true },
    { name: 'Andrew Muma',      store: 'north', division: 'M', roleType: 'tech',    fte: 1.0, pin: '', empNum: '029',  payRate: 37,    payType: 'Hourly', active: true },
    { name: 'Don Raes',         store: 'north', division: 'M', roleType: 'tech',    fte: 1.0, pin: '', empNum: '045',  payRate: 40,    payType: 'Hourly', active: true },
    { name: 'Pat Schaffner',    store: 'north', division: 'M', roleType: 'tech',    fte: 0.5, pin: '', empNum: '',     payRate: 35,    payType: 'Hourly', active: true },
    { name: 'Nate',             store: 'north', division: 'M', roleType: 'tech',    fte: 1.0, pin: '', empNum: '053', effName: 'NATHAN VERBURG', active: false },
    // ---- North support ----
    { name: 'Paul Lopez',       store: 'north', division: 'M', roleType: 'support', fte: 1.0, pin: '', payRate: 25,    payType: 'Hourly', active: true, queue: 'IM' },
    { name: 'Zach Noble-Welch', store: 'north', division: 'M', roleType: 'support', fte: 1.0, pin: '', payRate: 25,    payType: 'Hourly', active: true, queue: 'IM' },
    { name: 'Travis Wolfenden', store: 'north', division: 'M', roleType: 'support', fte: 1.0, pin: '', payRate: 27,    payType: 'Hourly', active: true, queue: 'IM' },
    { name: 'Bryan Smith',      store: 'north', division: 'M', roleType: 'support', fte: 1.0, pin: '', payRate: 27.5,  payType: 'Hourly', active: true, queue: 'WM' },
    // ---- Warranty admin ----
    { name: 'Andy Van Bommel',  store: 'north', division: 'M', roleType: 'warranty', fte: 1.0, pin: '', payRate: 33.5,  payType: 'Hourly', active: true },
    // ---- Office admin ----
    { name: 'Mell Hogg',        store: 'south', division: 'S', roleType: 'admin',   fte: 1.0, pin: '', payRate: 58000, payType: 'Salary', active: true },
    // ---- Managers (also email-auth Tier 1) ----
    { name: 'Steve Hayes',      store: 'south', division: 'S', roleType: 'manager', fte: 1.0, pin: '', email: 'steve@hydeparkequipment.ca', payRate: 100000, payType: 'Salary', active: true },
    { name: 'Bill Denison',     store: 'north', division: 'M', roleType: 'manager', fte: 1.0, pin: '', email: 'bill@hydeparkequipment.ca',  payRate: 97000,  payType: 'Salary', active: true }
  ];

  /* ---------------------------------------------------------------
   * SUPPORT WO QUEUE ASSIGNMENTS
   * POS Document# prefix -> who works that queue.
   *   WS = South work orders, WM = North work orders
   *   IS = South invoices,    IM = North invoices
   * ------------------------------------------------------------- */
  var WO_ASSIGNMENTS = {
    WS: { label: 'South Work Orders', store: 'south', staff: ['Stew', 'Steve'] },
    WM: { label: 'North Work Orders', store: 'north', staff: ['Bill', 'Bryan'] },
    IS: { label: 'South Invoices',    store: 'south', staff: ['Amy', 'Matt', 'Thorin'] },
    IM: { label: 'North Invoices',    store: 'north', staff: ['Zach', 'Travis', 'Paul', 'Bill'] }
  };

  // Which queues a given support employee can see.
  function queuesForStaff(name) {
    var out = [];
    Object.keys(WO_ASSIGNMENTS).forEach(function (q) {
      if (WO_ASSIGNMENTS[q].staff.indexOf(name) !== -1) out.push(q);
    });
    return out;
  }

  /* ---------------------------------------------------------------
   * BRANDING
   * ------------------------------------------------------------- */
  var BRAND = {
    name: 'Hyde Park Equipment',
    short: 'HPE',
    orange: '#E8620A',
    dark: '#1c1c1c'
  };

  /* ---------------------------------------------------------------
   * EXPORT
   * ------------------------------------------------------------- */
  global.HPE_CONFIG = {
    BACKEND: BACKEND,
    STORES: STORES,
    storeByDivision: storeByDivision,
    EMAIL_USERS: EMAIL_USERS,
    userForEmail: userForEmail,
    KPI_CONFIG: KPI_CONFIG,
    MANAGER_BONUS: MANAGER_BONUS,
    DEFAULT_STAFF: DEFAULT_STAFF,
    WO_ASSIGNMENTS: WO_ASSIGNMENTS,
    queuesForStaff: queuesForStaff,
    BRAND: BRAND
  };

})(typeof window !== 'undefined' ? window : this);
