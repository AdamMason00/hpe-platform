/* =====================================================================
 * HPE Service Platform — Google Apps Script backend
 * ---------------------------------------------------------------------
 * Deployed as a Web App ("Execute as me", "Anyone with the link").
 * The GitHub Pages frontend calls it via:
 *   - JSONP for GET  (doGet wraps the response in callback(...))
 *   - text/plain POST for writes (doPost reads e.postData.contents)
 *
 * Storage model:
 *   AppData sheet — one row per key (KPI, Staff, Efficiency, POS,
 *   Exclusions, Warranty) holding a JSON blob in column B. Simple,
 *   atomic, and easy to back up.
 *
 * Run initialSetup() once from the editor to create the Drive folder
 * tree and all sheet tabs.
 * ===================================================================== */

var SHEET_ID      = '1Ljh-Ycf1ut6TyV2NgXFRrypRUJzLRPpdUjl_yOIHrBw';
var DRIVE_FOLDER  = '1m9wv8eaWhAaLe1qZ0T0P35zUmt4NrSPQ';   // uploads + backups root
var PHOTOS_FOLDER = '1kbsKqfQp-Ms4YqwOtWxD2p3JYTiPEUu5';   // claim photos

var APPDATA_SHEET   = 'AppData';
var CHANGELOG_SHEET = 'ChangeLog';
var WOHIST_SHEET    = 'WO_Exclusion_History';
var BACKUP_KEEP     = 30;

/* ============================ ROUTING ============================ */
function doGet(e) {
  var params   = (e && e.parameters) || {};
  var action   = ((e && e.parameter && e.parameter.action) || 'ping');
  var callback = (e && e.parameter && e.parameter.callback) || '';
  var out;
  try {
    out = handleGet(action, e);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return reply(out, callback);
}

function doPost(e) {
  var out;
  try {
    var body    = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action  = body.action;
    var payload = body.payload || {};
    out = handlePost(action, payload);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return reply(out, '');
}

// JSON or JSONP reply.
function reply(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================ GET HANDLERS ============================ */
function handleGet(action, e) {
  switch (action) {
    case 'ping':           return { ok: true, pong: true, time: new Date().toISOString() };
    case 'getUser':        return getUser();
    case 'loadKPI':        return { ok: true, kpi: readBlob('KPI', {}) };
    case 'loadWarranty':   return { ok: true, warranty: readBlob('Warranty', {}) };
    case 'loadPOS':        return { ok: true, rows: readBlob('POS', { rows: [] }).rows || [], uploadedAt: readBlob('POS', {}).uploadedAt || null };
    case 'loadEfficiency': return { ok: true, rows: readBlob('Efficiency', { rows: [] }).rows || [], uploadedAt: readBlob('Efficiency', {}).uploadedAt || null };
    case 'loadExclusions': return { ok: true, exclusions: readBlob('Exclusions', {}) };
    case 'loadStaff':      return { ok: true, staff: readBlob('Staff', { staff: [] }).staff || [] };
    default:               return { ok: false, error: 'Unknown GET action: ' + action };
  }
}

function getUser() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  if (!email) { try { email = Session.getEffectiveUser().getEmail() || ''; } catch (e) {} }
  return { ok: true, email: email };
}

/* ============================ POST HANDLERS ============================ */
function handlePost(action, payload) {
  switch (action) {
    case 'saveKPI':        return saveBlob('KPI', payload.kpi || payload, 'kpi');
    case 'saveWarranty':   return saveBlob('Warranty', payload.warranty || payload, 'warranty');
    case 'savePOS':        return saveBlob('POS', { rows: payload.rows || [], uploadedAt: payload.uploadedAt || stamp() }, 'pos');
    case 'saveEfficiency': return saveBlob('Efficiency', { rows: payload.rows || [], uploadedAt: payload.uploadedAt || stamp() }, 'efficiency');
    case 'saveExclusions': return saveExclusions(payload.exclusions || {});
    case 'saveStaff':      return saveBlob('Staff', { staff: payload.staff || [] }, 'staff');
    case 'savePINs':       return savePINs(payload.pins || []);
    case 'uploadFile':     return uploadFile(payload);
    case 'backup':         return backupModule(payload.module || 'manual', payload.data || {});
    case 'importCSV':      return importCSV(payload);
    default:               return { ok: false, error: 'Unknown POST action: ' + action };
  }
}

/* ============================ BLOB STORAGE ============================ */
function appDataSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(APPDATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(APPDATA_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'UpdatedAt']]);
  }
  return sh;
}
function findKeyRow(sh, key) {
  var values = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues();
  for (var i = 1; i < values.length; i++) if (values[i][0] === key) return i + 1;
  return -1;
}
function readBlob(key, fallback) {
  var sh = appDataSheet();
  var row = findKeyRow(sh, key);
  if (row === -1) return fallback;
  var raw = sh.getRange(row, 2).getValue();
  if (!raw) return fallback;
  try { return JSON.parse(raw); } catch (e) { return fallback; }
}
function writeBlob(key, obj) {
  var sh = appDataSheet();
  var row = findKeyRow(sh, key);
  var json = JSON.stringify(obj);
  if (row === -1) {
    sh.appendRow([key, json, stamp()]);
  } else {
    sh.getRange(row, 2).setValue(json);
    sh.getRange(row, 3).setValue(stamp());
  }
}
function saveBlob(key, obj, module) {
  writeBlob(key, obj);
  logChange(module, 'save', key);
  backupModule(module, obj);
  return { ok: true, savedAt: stamp() };
}

/* ---- staff PIN merge (keeps other staff fields intact) ---- */
function savePINs(pins) {
  var data = readBlob('Staff', { staff: [] });
  var staff = data.staff || [];
  var byName = {};
  staff.forEach(function (s) { byName[s.name] = s; });
  pins.forEach(function (p) {
    if (byName[p.name]) byName[p.name].pin = String(p.pin || '');
  });
  writeBlob('Staff', { staff: staff });
  logChange('staff', 'savePINs', pins.length + ' pins');
  backupModule('staff', { staff: staff });
  return { ok: true, savedAt: stamp() };
}

/* ---- exclusions + history log ---- */
function saveExclusions(exclusions) {
  var prev = readBlob('Exclusions', {});
  writeBlob('Exclusions', exclusions);
  // Append any changes to the WO exclusion history sheet.
  var sh = sheetOrCreate(WOHIST_SHEET, ['Timestamp', 'WO#', 'Excluded', 'Note', 'By']);
  var by = currentUserEmail();
  Object.keys(exclusions).forEach(function (wo) {
    var cur = exclusions[wo] || {};
    var old = prev[wo] || {};
    if (JSON.stringify(cur) !== JSON.stringify(old)) {
      sh.appendRow([stamp(), wo, cur.excluded ? 'YES' : 'no', cur.note || '', by]);
    }
  });
  logChange('exclusions', 'save', Object.keys(exclusions).length + ' WOs');
  backupModule('exclusions', exclusions);
  return { ok: true, savedAt: stamp() };
}

/* ============================ FILE UPLOAD ============================ */
// payload: { filename, mimeType, data(base64), category(Efficiency|POS|Warranty), period }
function uploadFile(payload) {
  var root = DriveApp.getFolderById(DRIVE_FOLDER);
  var uploads = childFolder(root, 'Data Uploads');
  var periodFolder = childFolder(uploads, payload.period || 'FYTD');
  var catFolder = childFolder(periodFolder, payload.category || 'Misc');
  var bytes = Utilities.base64Decode(payload.data || '');
  var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', payload.filename || ('upload-' + stamp() + '.bin'));
  var file = catFolder.createFile(blob);
  logChange('upload', payload.category || 'file', file.getName());
  return { ok: true, fileId: file.getId(), url: file.getUrl(), savedAt: stamp() };
}

/* ============================ BACKUPS ============================ */
// Every save writes a timestamped JSON to Drive/Backups/{module}/, keep last 30.
function backupModule(module, data) {
  try {
    var root = DriveApp.getFolderById(DRIVE_FOLDER);
    var backups = childFolder(root, 'Backups');
    var modFolder = childFolder(backups, module || 'misc');
    var name = module + '-' + stamp().replace(/[:.]/g, '-') + '.json';
    modFolder.createFile(name, JSON.stringify(data, null, 2), 'application/json');
    pruneFolder(modFolder, BACKUP_KEEP);
  } catch (e) { /* backup failure must never block a save */ }
  return { ok: true };
}
function pruneFolder(folder, keep) {
  var files = [];
  var it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (var i = keep; i < files.length; i++) {
    try { files[i].setTrashed(true); } catch (e) {}
  }
}

/* ============================ CHANGE LOG ============================ */
function logChange(module, action, detail) {
  try {
    var sh = sheetOrCreate(CHANGELOG_SHEET, ['Timestamp', 'Module', 'Action', 'Detail', 'User']);
    sh.appendRow([stamp(), module || '', action || '', detail || '', currentUserEmail()]);
  } catch (e) {}
}

/* ============================ CSV IMPORT ============================ */
// payload: { key:'KPI'|'Staff'|..., rows:[...] } — stores parsed rows as a blob.
function importCSV(payload) {
  var key = payload.key || 'Import';
  writeBlob(key, payload.rows || payload.data || []);
  logChange('import', key, (payload.rows ? payload.rows.length : 0) + ' rows');
  return { ok: true, savedAt: stamp() };
}

/* ============================ HELPERS ============================ */
function stamp() { return new Date().toISOString(); }
function currentUserEmail() {
  try { return Session.getActiveUser().getEmail() || 'system'; } catch (e) { return 'system'; }
}
function childFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function sheetOrCreate(name, headers) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ============================ INITIAL SETUP ============================ */
// Run once from the Apps Script editor (authorize when prompted).
function initialSetup() {
  // 1. Sheet tabs
  appDataSheet();
  sheetOrCreate(CHANGELOG_SHEET, ['Timestamp', 'Module', 'Action', 'Detail', 'User']);
  sheetOrCreate(WOHIST_SHEET, ['Timestamp', 'WO#', 'Excluded', 'Note', 'By']);

  // 2. Seed empty blobs if missing
  ['KPI', 'Warranty', 'POS', 'Efficiency', 'Exclusions', 'Staff'].forEach(function (k) {
    if (findKeyRow(appDataSheet(), k) === -1) {
      writeBlob(k, k === 'Staff' ? { staff: [] } : (k === 'POS' || k === 'Efficiency') ? { rows: [] } : {});
    }
  });

  // 3. Drive folder tree
  var root = DriveApp.getFolderById(DRIVE_FOLDER);
  var uploads = childFolder(root, 'Data Uploads');
  ['Q1', 'Q2', 'Q3', 'Q4', 'FYTD'].forEach(function (p) {
    var pf = childFolder(uploads, p);
    ['Efficiency', 'POS', 'Warranty'].forEach(function (c) { childFolder(pf, c); });
  });
  var backups = childFolder(root, 'Backups');
  ['kpi', 'warranty', 'pos', 'efficiency', 'exclusions', 'staff'].forEach(function (m) { childFolder(backups, m); });

  Logger.log('initialSetup complete.');
  return 'ok';
}

/* ============================ LEGACY MIGRATION ============================
 * One-time importer: copies the existing "KPI Manager 2026" data (a single
 * AppData JSON blob, schema _version 5.0, plus the per-employee efficiency
 * blob) into the new multi-key structure used by this backend.
 *
 * It does NOT modify the legacy tabs — it only reads them and writes the
 * new KPI / Staff / Efficiency blobs. Safe to re-run (idempotent overwrite).
 * Run migrateFromLegacy() once from the editor, check the log summary.
 * ======================================================================= */
var MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function migrateFromLegacy() {
  var legacy = findLegacyBlob_(function (o) { return o && o._version && Array.isArray(o.employees); });
  if (!legacy) throw new Error('Legacy AppData blob (with _version & employees) not found in the spreadsheet.');
  var effBlob = findLegacyBlob_(function (o) {
    if (!o || typeof o !== 'object' || Array.isArray(o)) return false;
    var k = Object.keys(o)[0];
    return k && o[k] && o[k].months && (o[k].empNum !== undefined || o[k].name !== undefined);
  });

  var kpi = mapLegacyKPI_(legacy);
  var staff = mapLegacyStaff_(legacy.employees);
  var eff = effBlob ? mapLegacyEfficiency_(effBlob) : { rows: [], uploadedAt: stamp() };

  writeBlob('KPI', kpi);
  writeBlob('Staff', { staff: staff });
  writeBlob('Efficiency', eff);
  logChange('migration', 'migrateFromLegacy',
    staff.length + ' staff, ' + Object.keys(kpi.quarters).length + ' quarters, ' + eff.rows.length + ' eff rows');

  var summary = {
    staff: staff.length,
    quarters: Object.keys(kpi.quarters),
    managers: Object.keys(kpi.managerBonus),
    effRows: eff.rows.length,
    config: kpi.config
  };
  Logger.log('migrateFromLegacy summary: ' + JSON.stringify(summary, null, 2));
  return summary;
}

// Scan every sheet/cell, JSON.parse candidates, return the first that passes `test`.
function findLegacyBlob_(test) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var rng = sheets[s].getDataRange().getValues();
    for (var r = 0; r < rng.length; r++) {
      for (var c = 0; c < rng[r].length; c++) {
        var v = rng[r][c];
        if (typeof v !== 'string' || v.length < 20 || v.charAt(0) !== '{') continue;
        var obj;
        try { obj = JSON.parse(v); } catch (e) { continue; }
        if (test(obj)) return obj;
      }
    }
  }
  return null;
}

function legacyStore_(s) {
  s = String(s || '').toLowerCase();
  if (s.indexOf('south') === 0) return 'south';
  if (s.indexOf('north') === 0) return 'north';
  return 'north'; // 'Both' (e.g. warranty admin) → north
}
function legacyDivision_(store) { return store === 'south' ? 'S' : 'M'; }
function legacyRole_(role) {
  role = String(role || '').toLowerCase();
  if (role.indexOf('manager') !== -1) return 'manager';
  if (role.indexOf('support') !== -1) return 'support';
  if (role.indexOf('warranty') !== -1 || role.indexOf('admin') !== -1) return 'admin';
  return 'tech';
}

function mapLegacyStaff_(employees) {
  return (employees || []).map(function (e) {
    var store = legacyStore_(e.store);
    return {
      id: e.id,
      name: e.name,
      store: store,
      division: legacyDivision_(store),
      roleType: legacyRole_(e.role),
      fte: e.fte == null ? 1 : Number(e.fte),
      pin: e.pin == null ? '' : String(e.pin),
      payRate: e.payRate,
      payType: e.payType,
      payHistory: e.payHistory || [],
      active: e.active !== false,
      queue: e.queue || ''
    };
  });
}

function mapLegacyKPI_(legacy) {
  var cfg = legacy.config || {};
  var config = {
    techEff: numOr_(cfg.techEff, 75), comeback: numOr_(cfg.comeback, 2),
    svcGm: numOr_(cfg.svcGm, 78), partsGm: numOr_(cfg.partsGm, 32),
    wipMax: numOr_(cfg.wipMax, 2), cap: numOr_(cfg.cap, 9000),
    techShare: numOr_(cfg.techShare, 73), growthRate: numOr_(cfg.growthRate, 30),
    warrantyAnnual: numOr_(legacy.warrantyAnnual, 5000), topUpPool: numOr_(legacy.topUpPool, 50000),
    woWarnDays: 30, woCriticalDays: 60, effHighFlag: 100, effLowFlag: 75
  };

  // quarters: legacy {Q1..Q4}{North,South} -> '2026-Qn' { south, north }
  var quarters = {};
  var lq = legacy.quarters || {};
  ['Q1', 'Q2', 'Q3', 'Q4'].forEach(function (q) {
    var key = '2026-' + q;
    var src = lq[q];
    if (!src) return;
    var out = {};
    ['South', 'North'].forEach(function (st) {
      var m = src[st];
      if (!m) return;
      out[st.toLowerCase()] = {
        hrs: numOr_(m.hrs, 0), billed: numOr_(m.billed, 0), comeback: numOr_(m.comeback, 0),
        svcGm: numOr_(m.svcGm, 0), partsGm: numOr_(m.partsGm, 0),
        svcRev: numOr_(m.svcRev, 0), endingWip: numOr_(m.endingWip, 0),
        notes: m.notes || '', paid: !!m.paid, paidDate: m.paidDate || ''
      };
    });
    if (Object.keys(out).length) quarters[key] = out;
  });

  // manager KPI bonuses: legacy {South,North}{cap,paid{Qn}} -> email-keyed
  var managerBonus = {};
  var mk = legacy.mgrKPIBonuses || {};
  var emailByStore = { south: 'steve@hydeparkequipment.ca', north: 'bill@hydeparkequipment.ca' };
  ['South', 'North'].forEach(function (st) {
    var src = mk[st];
    if (!src) return;
    var email = emailByStore[st.toLowerCase()];
    var paid = {};
    Object.keys(src.paid || {}).forEach(function (q) { paid['2026-' + q] = !!src.paid[q]; });
    managerBonus[email] = { annualCap: numOr_(src.cap, 6000), paid: paid };
  });

  return {
    config: config,
    quarters: quarters,
    managerBonus: managerBonus,
    payments: [],
    growth: { rate: config.growthRate, bank: { south: 0, north: 0 }, paidOut: 0, history: [] },
    legacy: { mgrBonuses: legacy.mgrBonuses || {}, warrantyQuarters: legacy.warrantyQuarters || {},
              topUpScores: legacy.topUpScores || {}, prevYear: legacy.prevYear || {} }
  };
}

function mapLegacyEfficiency_(effBlob) {
  var rows = [];
  Object.keys(effBlob).forEach(function (key) {
    var e = effBlob[key];
    if (!e || !e.months) return;
    var div = String(e.division || key.split('|').pop() || '').toUpperCase().charAt(0);
    var name = e.name || key.split('|')[1] || key;
    Object.keys(e.months).forEach(function (mk) {
      var m = e.months[mk];
      var reported = numOr_(m.hrRep, 0), billed = numOr_(m.hrBil, 0);
      if (reported <= 0 && billed <= 0) return;
      var monthNum = parseInt(String(mk).slice(4, 6), 10);
      rows.push({
        month: MONTH_ABBR[monthNum - 1] || String(mk),
        name: name, division: div, docNum: '',
        reported: reported, billed: billed,
        eff: reported > 0 ? (billed / reported * 100) : 0
      });
    });
  });
  return { rows: rows, uploadedAt: stamp() };
}

function numOr_(v, d) { var n = parseFloat(v); return isNaN(n) ? d : n; }

/* Quick self-test from the editor. */
function testPing() { Logger.log(JSON.stringify(handleGet('ping', {}))); }
