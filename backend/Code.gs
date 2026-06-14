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

/* Quick self-test from the editor. */
function testPing() { Logger.log(JSON.stringify(handleGet('ping', {}))); }
