/* =====================================================================
 * HPE KPI Incentive Manager — application logic
 * ---------------------------------------------------------------------
 * Sections:
 *   1. Boot / auth / state
 *   2. Small helpers (DOM, format, toast)
 *   3. Periods
 *   4. Persistence (load/save to backend)
 *   5. KPI scoring + bonus maths
 *   6. Navigation + routing (role-aware)
 *   7. Page renderers (overview, data entry, people, tools)
 *   8. Role dashboards (store / tech / support)
 *   9. Configuration + consolidated .xlsx uploads
 *  10. Manager KPI PDF export
 * ===================================================================== */
(function () {
'use strict';

var CFG  = window.HPE_CONFIG;
var API  = window.HPE_API;
var AUTH = window.HPE_AUTH;

/* ============================ 1. BOOT / STATE ============================ */
var SESSION = AUTH.requireSession('../index.html');
if (!SESSION) return;            // redirected to login

var STATE = {
  kpi: {
    config: Object.assign({}, CFG.KPI_CONFIG),
    quarters: {},          // key -> { south:{...metrics}, north:{...metrics} }
    managerBonus: {},      // email -> { annualCap, paid:{periodKey:true} }
    payments: [],          // [{date, employee, period, amount, note}]
    growth: { rate: CFG.KPI_CONFIG.growthRate, threshold: 4, bank: { south: 0, north: 0 }, paidOut: 0, history: [], allocation: {} },
    effExclusions: {},     // non-billable entry key -> { excluded:bool, note:'' }
    manualBonuses: {},     // period -> { employeeName -> { amount, note } }
    warrantyAdmin: { bonusAmount: CFG.KPI_CONFIG.warrantyAnnual || 5000, exclusions: {}, paid: {} }
  },
  staff: [],
  // efficiency is stored AGGREGATED (raw labour files have thousands of rows
  // that won't fit a single Sheets cell): per-tech/month totals, a non-billable
  // detail list (for review/exclude), and only the flagged WOs.
  efficiency: { byTechMonth: {}, techDiv: {}, nonBillable: [], flagged: [], uploadedAt: null },
  nbDetail: [],   // in-memory only (this session's upload): individual non-billable punches for drill-down
  pos: { rows: [], uploadedAt: null },
  warrantyWO: { rows: [], uploadedAt: null },   // parsed warranty work orders (KPI source)
  warrantyUpload: { uploadedAt: null },
  exclusions: {}           // docNum -> { excluded:bool, note:'' }
};

var CURRENT_PERIOD = null;   // e.g. '2026-Q2'
var CURRENT_PAGE   = null;

/* ============================ 2. HELPERS ============================ */
function $(s, root){ return (root || document).querySelector(s); }
function el(tag, attrs, html){
  var n = document.createElement(tag);
  if (attrs) Object.keys(attrs).forEach(function(k){
    if (k === 'class') n.className = attrs[k];
    else if (k === 'html') n.innerHTML = attrs[k];
    else n.setAttribute(k, attrs[k]);
  });
  if (html != null) n.innerHTML = html;
  return n;
}
function money(n){
  n = Number(n) || 0;
  return '$' + n.toLocaleString('en-CA', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}
function money2(n){
  n = Number(n) || 0;
  return '$' + n.toLocaleString('en-CA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function pct(n, d){ if (n == null || isNaN(n)) return '—'; return (Number(n)).toFixed(d == null ? 1 : d) + '%'; }
function num(v){ var n = parseFloat(String(v).replace(/[$,%\s]/g,'')); return isNaN(n) ? 0 : n; }
function esc(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, function(c){
  return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]; }); }

function toast(msg, kind){
  var t = el('div', { class: 'toast ' + (kind || '') }, esc(msg));
  $('#toast').appendChild(t);
  setTimeout(function(){ t.style.opacity = '0'; t.style.transition = '.4s'; }, 3200);
  setTimeout(function(){ if (t.parentNode) t.parentNode.removeChild(t); }, 3700);
}

function storeName(id){ return id === 'south' ? 'South Store' : id === 'north' ? 'North Store' : id; }
function storeDivision(id){ return id === 'south' ? 'S' : 'M'; }
function divisionStore(div){ var s = CFG.storeByDivision(div); return s ? s.id : null; }

/* ============================ 3. PERIODS ============================ */
// Fiscal year inferred from today; we expose the 4 quarters of the active FY.
function buildPeriods(){
  var year = 2026;                       // current FY anchor (today 2026-06-13)
  var list = [];
  for (var q = 1; q <= 4; q++) list.push({ key: year + '-Q' + q, label: 'FY' + year + ' · Q' + q, year: year, q: q });
  return list;
}
var PERIODS = buildPeriods();

function periodLabel(key){
  for (var i = 0; i < PERIODS.length; i++) if (PERIODS[i].key === key) return PERIODS[i].label;
  return key;
}
function periodsInYear(year){
  return PERIODS.filter(function(p){ return p.year === year; });
}

/* ============================ 4. PERSISTENCE ============================ */
function loadAll(){
  setBusy(true);
  var jobs = [
    API.loadKPI().then(function(r){ mergeKPI(r); }).catch(noop),
    API.loadStaff().then(function(r){ STATE.staff = normaliseStaff((r && (r.staff || r.data)) || []); }).catch(function(){ STATE.staff = normaliseStaff(CFG.DEFAULT_STAFF); }),
    API.loadEfficiency().then(function(r){ ingestEfficiency(r, r && r.uploadedAt); }).catch(noop),
    API.loadPOS().then(function(r){ if (r && (r.rows || r.data)) ingestPOS(r.rows || r.data, r.uploadedAt); }).catch(noop),
    API.loadExclusions().then(function(r){ STATE.exclusions = (r && (r.exclusions || r.data)) || {}; }).catch(noop),
    API.loadWarranty().then(function(r){ var w = r && (r.warranty || r.data); if (w && w.wo) STATE.warrantyWO = { rows: w.wo, uploadedAt: w.uploadedAt || null }; }).catch(noop)
  ];
  return Promise.all(jobs).then(function(){
    if (!STATE.staff || !STATE.staff.length) STATE.staff = normaliseStaff(CFG.DEFAULT_STAFF);
    ensureManagerBonus();
    rebuildEfficiencyIndex();   // link Labour-data names to roster after both load
    setBusy(false);
  }).catch(function(e){ setBusy(false); toast('Load issue: ' + e.message, 'bad'); });
}
function noop(){}

function mergeKPI(r){
  var d = (r && (r.kpi || r.data)) || r;
  if (!d || typeof d !== 'object') return;
  if (d.config)        STATE.kpi.config = Object.assign({}, CFG.KPI_CONFIG, d.config);
  if (d.quarters)      STATE.kpi.quarters = d.quarters;
  if (d.managerBonus)  STATE.kpi.managerBonus = d.managerBonus;
  if (d.payments)      STATE.kpi.payments = d.payments;
  if (d.growth)        STATE.kpi.growth = Object.assign(STATE.kpi.growth, d.growth);
  if (d.effExclusions) STATE.kpi.effExclusions = d.effExclusions;
  if (d.manualBonuses) STATE.kpi.manualBonuses = d.manualBonuses;
  if (d.warrantyAdmin) STATE.kpi.warrantyAdmin = Object.assign(STATE.kpi.warrantyAdmin, d.warrantyAdmin);
}
function normaliseStaff(arr){
  return (arr || []).map(function(s){
    return {
      id: s.id, name: s.name, store: s.store, division: s.division || storeDivision(s.store),
      roleType: s.roleType || s.role || 'tech', fte: s.fte == null ? 1 : Number(s.fte),
      pin: s.pin == null ? '' : String(s.pin), queue: s.queue || '',
      // pay / compensation fields (preserved from backend + migration)
      payRate: s.payRate == null ? 0 : Number(s.payRate),
      payType: s.payType || 'Hourly',
      vacationWeeks: s.vacationWeeks == null ? 0 : Number(s.vacationWeeks),
      payHistory: Array.isArray(s.payHistory) ? s.payHistory : [],
      active: s.active !== false,
      effName: s.effName || '',  // Labour-data name override for efficiency linkage
      empNum: s.empNum == null ? '' : String(s.empNum).trim(),
      email: s.email || ''
    };
  });
}

/* ---- compensation helpers ----
 * Approximate annual income. Hourly assumes a standard 40h week × 52 weeks
 * (paid vacation included), scaled by FTE. Salary is taken as-is. */
var HOURS_PER_WEEK = 40, WEEKS_PER_YEAR = 52;
function annualIncome(s){
  var rate = num(s.payRate), fte = (s.fte == null ? 1 : Number(s.fte));
  if ((s.payType || 'Hourly') === 'Salary') return rate;        // salary already annual
  return rate * HOURS_PER_WEEK * WEEKS_PER_YEAR * fte;
}
function wageLabel(s){
  return (s.payType || 'Hourly') === 'Salary'
    ? money(s.payRate) + '/yr'
    : money2(s.payRate) + '/hr' + (s.fte && s.fte !== 1 ? ' · ' + s.fte + ' FTE' : '');
}
function ensureManagerBonus(){
  Object.keys(CFG.MANAGER_BONUS).forEach(function(email){
    if (!STATE.kpi.managerBonus[email]) {
      STATE.kpi.managerBonus[email] = { annualCap: CFG.MANAGER_BONUS[email].annualCap, paid: {} };
    }
    if (!STATE.kpi.managerBonus[email].paid) STATE.kpi.managerBonus[email].paid = {};
  });
}

function saveKPI(){
  return API.saveKPI({ kpi: STATE.kpi }).then(function(){ toast('Saved', 'ok'); })
    .catch(function(e){ toast('Save failed: ' + e.message, 'bad'); });
}
function saveStaff(){
  return API.saveStaff({ staff: STATE.staff })
    .then(function(){ toast('Roster saved', 'ok'); })
    .catch(function(e){ toast('Save failed: ' + e.message, 'bad'); });
}
function savePINs(){
  var pins = STATE.staff.map(function(s){ return { name: s.name, pin: s.pin }; });
  return API.savePINs({ pins: pins }).then(function(){ toast('PINs saved', 'ok'); })
    .catch(function(e){ toast('Save failed: ' + e.message, 'bad'); });
}
function saveExclusions(){
  return API.saveExclusions({ exclusions: STATE.exclusions })
    .then(function(){ toast('Exclusions saved', 'ok'); })
    .catch(function(e){ toast('Save failed: ' + e.message, 'bad'); });
}

function setBusy(b){
  var rb = $('#reloadBtn'); if (!rb) return;
  rb.innerHTML = b ? '<span class="spinner dark"></span>' : '⟳';
  rb.disabled = !!b;
}

/* ============================ 5. KPI SCORING + BONUS ============================ */
// metrics: { hrs, billed, comeback, svcGm, partsGm, svcRev, endingWip }
function techEffOf(m){ return (m && m.hrs > 0) ? (m.billed / m.hrs * 100) : 0; }
function wipPctOf(m){ return (m && m.svcRev > 0) ? (m.endingWip / m.svcRev * 100) : 0; }

function scoreKPIs(m){
  var c = STATE.kpi.config;
  m = m || {};
  var eff = techEffOf(m);
  var kpi1 = (eff >= c.techEff) && (num(m.comeback) <= c.comeback);
  var kpi2 = num(m.svcGm)   >= c.svcGm;
  var kpi3 = num(m.partsGm) >= c.partsGm;
  var kpi4 = wipPctOf(m)    <= c.wipMax;
  var points = (kpi1?1:0) + (kpi2?1:0) + (kpi3?1:0) + (kpi4?1:0);
  return { kpi1: kpi1, kpi2: kpi2, kpi3: kpi3, kpi4: kpi4, points: points,
           eff: eff, wipPct: wipPctOf(m) };
}

function metricsFor(periodKey, store){
  var q = STATE.kpi.quarters[periodKey];
  return (q && q[store]) || null;
}

// Manager KPI bonus for a period (cap ÷ 4 quarters ÷ 4 KPIs × hits).
function managerBonusFor(email, periodKey){
  var mb = STATE.kpi.managerBonus[email];
  var store = CFG.MANAGER_BONUS[email] ? CFG.MANAGER_BONUS[email].store : null;
  var cap = mb ? mb.annualCap : 0;
  var quarterlyMax = cap / 4;
  var perKpi = quarterlyMax / 4;
  var m = metricsFor(periodKey, store);
  var score = m ? scoreKPIs(m) : { points: 0, kpi1:false,kpi2:false,kpi3:false,kpi4:false };
  var payout = score.points * perKpi;
  return {
    email: email, store: store, cap: cap, quarterlyMax: quarterlyMax, perKpi: perKpi,
    score: score, payout: payout, paid: !!(mb && mb.paid && mb.paid[periodKey])
  };
}
function managerYtd(email){
  var earned = 0, paid = 0;
  periodsInYear(2026).forEach(function(p){
    var b = managerBonusFor(email, p.key);
    earned += b.payout;
    if (b.paid) paid += b.payout;
  });
  var cap = STATE.kpi.managerBonus[email] ? STATE.kpi.managerBonus[email].annualCap : 0;
  return { earned: earned, paid: paid, cap: cap, remaining: Math.max(0, cap - earned) };
}

// Tech/support bonus pool for a store-quarter. Paid only if store scores ≥2/4.
function techBonusPool(periodKey, store){
  var c = STATE.kpi.config;
  var m = metricsFor(periodKey, store);
  var score = m ? scoreKPIs(m) : { points: 0 };
  var eligible = score.points >= 2;
  var pool = eligible ? c.cap : 0;
  var techPool = pool * (c.techShare / 100);
  var supportPool = pool * (1 - c.techShare / 100);
  var roster = STATE.staff.filter(function(s){ return s.store === store; });
  var techs = roster.filter(function(s){ return s.roleType === 'tech'; });
  var support = roster.filter(function(s){ return s.roleType === 'support'; });
  var techFte = techs.reduce(function(a,s){ return a + (s.fte||0); }, 0) || 1;
  var supFte  = support.reduce(function(a,s){ return a + (s.fte||0); }, 0) || 1;
  var lines = [];
  techs.forEach(function(s){ lines.push({ name: s.name, group: 'Tech', fte: s.fte, amount: techPool * (s.fte / techFte) }); });
  support.forEach(function(s){ lines.push({ name: s.name, group: 'Support', fte: s.fte, amount: supportPool * (s.fte / supFte) }); });
  return { eligible: eligible, points: score.points, pool: pool, techPool: techPool, supportPool: supportPool, lines: lines };
}

/* ---- Growth bonus accrual (per store, per quarter) ----
 * If combined parts+service income grows >= threshold% vs the prior-year same
 * quarter, accrue rate% (30%) of that store's tech+support pool to its bank. */
function growthFromData(store, m){
  m = m || {};
  var curParts = num(m.curParts), curSvc = num(m.curSvc), priorParts = num(m.priorParts), priorSvc = num(m.priorSvc);
  var curTotal = curParts + curSvc, priorTotal = priorParts + priorSvc;
  var pctOf = function(cur, prior){ return prior > 0 ? (cur - prior) / prior * 100 : null; };
  var rate = num(STATE.kpi.growth.rate) || 30;
  var threshold = (STATE.kpi.growth.threshold != null) ? num(STATE.kpi.growth.threshold) : 4;
  var combG = pctOf(curTotal, priorTotal);
  var gate = combG != null && combG >= threshold;
  return {
    curParts: curParts, curSvc: curSvc, priorParts: priorParts, priorSvc: priorSvc,
    curTotal: curTotal, priorTotal: priorTotal,
    partsG: pctOf(curParts, priorParts), svcG: pctOf(curSvc, priorSvc), combG: combG,
    gate: gate, rate: rate, threshold: threshold
  };
}
function storeGrowth(periodKey, store){
  var g = growthFromData(store, metricsFor(periodKey, store) || {});
  g.pool = techBonusPool(periodKey, store).pool;
  g.accrual = g.gate ? (g.rate / 100) * g.pool : 0;
  return g;
}
function storeBank(store){
  var t = 0; periodsInYear(2026).forEach(function(p){ t += storeGrowth(p.key, store).accrual; });
  return t;
}
function combinedGrowthPool(){
  return storeBank('south') + storeBank('north');
}

/* ============================ 6. NAV / ROUTING ============================ */
var NAV = [
  { group: 'Overview', items: [
    { id: 'dashboard',  label: 'Dashboard',           ic: '▦', roles: ['admin','manager'] },
    { id: 'scoreboard', label: 'Quarterly Scoreboard', ic: '▣', roles: ['admin','manager'] },
    { id: 'annual',     label: 'Annual Summary',       ic: '∑', roles: ['admin','manager'] }
  ]},
  { group: 'Data Entry', items: [
    { id: 'quarterly',  label: 'Quarterly Results',    ic: '✎', roles: ['admin','manager'] }
  ]},
  { group: 'Payroll & People', items: [
    { id: 'staff',      label: 'Staff Roster',         ic: '👥', roles: ['admin','manager'] },
    { id: 'payroll',    label: 'Payroll Summary',      ic: '💲', roles: ['admin','manager'] },
    { id: 'efficiency', label: 'Tech Efficiency',      ic: '⚙', roles: ['admin','manager'] },
    { id: 'payments',   label: 'Payment History',      ic: '🧾', roles: ['admin','manager'] },
    { id: 'growth',     label: 'Growth Bonus',         ic: '📈', roles: ['admin','manager'] },
    { id: 'manager-bonuses', label: 'Manager Bonuses', ic: '🏅', roles: ['admin','manager'] }
  ]},
  { group: 'Tools', items: [
    { id: 'scenario',   label: 'Scenario Testing',     ic: '🧪', roles: ['admin','manager'] },
    { id: 'warranty-kpi', label: 'Warranty Admin KPI', ic: '🛠️', roles: ['admin'] },
    { id: 'config',     label: 'Configuration',        ic: '⚒', roles: ['admin'] }
  ]},
  { group: 'My Dashboard', items: [
    { id: 'store',   label: 'Store Dashboard',   ic: '🏬', roles: ['manager'] },
    { id: 'tech',    label: 'My Efficiency',     ic: '📊', roles: ['tech'] },
    { id: 'support', label: 'My WO Queue',       ic: '📋', roles: ['support'] }
  ]}
];

function visibleItems(){
  var out = [];
  NAV.forEach(function(g){
    var items = g.items.filter(function(it){ return it.roles.indexOf(SESSION.role) !== -1; });
    if (items.length) out.push({ group: g.group, items: items });
  });
  return out;
}

function buildNav(){
  var nav = $('#nav'); nav.innerHTML = '';
  visibleItems().forEach(function(g){
    var grp = el('div', { class: 'nav-group' });
    grp.appendChild(el('div', { class: 'nav-label' }, esc(g.group)));
    g.items.forEach(function(it){
      var item = el('div', { class: 'nav-item', 'data-page': it.id },
        '<span class="ic">' + it.ic + '</span><span>' + esc(it.label) + '</span>');
      item.addEventListener('click', function(){ go(it.id); });
      grp.appendChild(item);
    });
    nav.appendChild(grp);
  });
}

var PAGE_META = {
  dashboard:  ['Dashboard', 'Quarterly performance overview'],
  scoreboard: ['Quarterly Scoreboard', 'Side-by-side store KPI results'],
  annual:     ['Annual Summary', 'Year-to-date totals and trends'],
  quarterly:  ['Quarterly Results', 'Enter quarterly data per store'],
  staff:      ['Staff Roster', 'Employees, roles, FTE & PINs'],
  payroll:    ['Payroll Summary', 'Quarterly payout breakdown'],
  efficiency: ['Tech Efficiency', 'Monthly efficiency from Labour data'],
  payments:   ['Payment History', 'Actual payments made'],
  growth:     ['Growth Bonus', 'Growth bank & year-end payout'],
  'manager-bonuses': ['Manager Bonuses', 'Steve & Bill KPI bonus'],
  scenario:   ['Scenario Testing', 'What-if calculator'],
  'warranty-kpi': ['Warranty Admin KPI', 'Close warranty WOs within 14 days of last punch'],
  config:     ['Configuration', 'Thresholds, uploads & staff config'],
  store:      ['Store Dashboard', 'Your store KPIs, efficiency & open WOs'],
  tech:       ['My Efficiency', 'Your monthly results & flagged WOs'],
  support:    ['My WO Queue', 'Your open work-order queue']
};

function go(pageId){
  CURRENT_PAGE = pageId;
  document.querySelectorAll('.nav-item').forEach(function(n){
    n.classList.toggle('active', n.getAttribute('data-page') === pageId);
  });
  document.querySelectorAll('.page').forEach(function(p){ p.classList.remove('active'); });
  var sec = $('#page-' + pageId);
  sec.classList.add('active');
  var meta = PAGE_META[pageId] || [pageId, ''];
  $('#pageTitle').textContent = meta[0];
  $('#pageSub').textContent = meta[1];
  $('#content').scrollTop = 0;
  (RENDER[pageId] || function(){ sec.innerHTML = '<div class="empty">Coming soon.</div>'; })(sec);
}

/* default landing page per role */
function landingPage(){
  if (SESSION.role === 'tech') return 'tech';
  if (SESSION.role === 'support') return 'support';
  return 'dashboard';
}

/* ============================ 7. PAGE RENDERERS ============================ */
var RENDER = {};

function kpiChips(score){
  var c = STATE.kpi.config;
  var defs = [
    ['Efficiency & Comeback', score.kpi1, '≥' + c.techEff + '% & ≤' + c.comeback + '%'],
    ['Service GM',            score.kpi2, '≥' + c.svcGm + '%'],
    ['Parts GM',              score.kpi3, '≥' + c.partsGm + '%'],
    ['Open WIP',              score.kpi4, '≤' + c.wipMax + '%']
  ];
  return '<div class="kpi-row">' + defs.map(function(d){
    return '<div class="kpi-chip ' + (d[1] ? 'hit' : 'miss') + '">' +
      (d[1] ? '✓' : '✕') + ' ' + esc(d[0]) + ' <span class="muted">' + d[2] + '</span></div>';
  }).join('') + '</div>';
}

/* ---- Dashboard ---- */
RENDER.dashboard = function(sec){
  var pk = CURRENT_PERIOD;
  var stores = managerStores();
  var html = '';

  // top stat row
  var totalPoints = 0, totalPool = 0;
  stores.forEach(function(st){
    var m = metricsFor(pk, st); var s = m ? scoreKPIs(m) : { points: 0 };
    totalPoints += s.points; totalPool += techBonusPool(pk, st).pool;
  });
  html += '<div class="grid g4">';
  html += stat('Period', periodLabel(pk), stores.length + ' store' + (stores.length>1?'s':''), '');
  html += stat('KPIs Hit', totalPoints + ' / ' + (stores.length * 4), 'across stores', 'blue');
  html += stat('Tech Bonus Pool', money(totalPool), 'this quarter', 'ok');
  var eff = avgEfficiency(pk, null);
  html += stat('Avg Tech Eff', eff == null ? '—' : pct(eff), 'target ' + STATE.kpi.config.techEff + '%', eff != null && eff >= STATE.kpi.config.techEff ? 'ok' : 'bad');
  html += '</div><div class="spacer"></div>';

  // per-store scorecards
  html += '<div class="grid ' + (stores.length > 1 ? 'g2' : '') + '">';
  stores.forEach(function(st){
    var m = metricsFor(pk, st);
    var s = m ? scoreKPIs(m) : null;
    html += '<div class="card">';
    html += '<div class="section-title"><h3>' + storeName(st) + '</h3>' +
            (s ? '<span class="pill ' + (s.points>=2?'ok':'bad') + '">' + s.points + '/4 KPIs</span>' : '<span class="pill muted">no data</span>') + '</div>';
    if (s) {
      html += kpiChips(s);
      html += '<div class="spacer"></div><div class="grid g2">';
      html += miniStat('Tech Efficiency', pct(s.eff), s.eff >= STATE.kpi.config.techEff);
      html += miniStat('Comeback', pct(m.comeback), num(m.comeback) <= STATE.kpi.config.comeback);
      html += miniStat('Service GM', pct(m.svcGm), num(m.svcGm) >= STATE.kpi.config.svcGm);
      html += miniStat('Parts GM', pct(m.partsGm), num(m.partsGm) >= STATE.kpi.config.partsGm);
      html += miniStat('Open WIP', pct(s.wipPct), s.wipPct <= STATE.kpi.config.wipMax);
      html += miniStat('Service Rev', money(m.svcRev), true);
      html += '</div>';
    } else {
      html += '<div class="empty">No quarterly results entered. Go to <b>Quarterly Results</b>.</div>';
    }
    html += '</div>';
  });
  html += '</div>';
  sec.innerHTML = html;
};

function miniStat(label, value, good){
  return '<div class="stat ' + (good ? 'ok' : 'bad') + '"><div class="label">' + esc(label) +
    '</div><div class="value" style="font-size:22px">' + value + '</div></div>';
}
function stat(label, value, meta, kind){
  return '<div class="stat ' + (kind||'') + '"><div class="label">' + esc(label) + '</div>' +
    '<div class="value">' + value + '</div>' + (meta ? '<div class="meta">' + esc(meta) + '</div>' : '') + '</div>';
}

/* ---- Quarterly Scoreboard ---- */
RENDER.scoreboard = function(sec){
  var pk = CURRENT_PERIOD;
  var stores = managerStores();
  var c = STATE.kpi.config;
  var rows = [
    ['Tech Efficiency', function(m,s){ return { v: pct(s.eff), ok: s.eff >= c.techEff }; }, '≥' + c.techEff + '%'],
    ['Comeback %',      function(m){ return { v: pct(m.comeback), ok: num(m.comeback) <= c.comeback }; }, '≤' + c.comeback + '%'],
    ['Service GM',      function(m){ return { v: pct(m.svcGm), ok: num(m.svcGm) >= c.svcGm }; }, '≥' + c.svcGm + '%'],
    ['Parts GM',        function(m){ return { v: pct(m.partsGm), ok: num(m.partsGm) >= c.partsGm }; }, '≥' + c.partsGm + '%'],
    ['Open WIP %',      function(m,s){ return { v: pct(s.wipPct), ok: s.wipPct <= c.wipMax }; }, '≤' + c.wipMax + '%'],
    ['Service Revenue', function(m){ return { v: money(m.svcRev), ok: null }; }, ''],
    ['Ending WIP $',    function(m){ return { v: money(m.endingWip), ok: null }; }, '']
  ];
  var html = '<div class="card"><div class="section-title"><h3>Scoreboard — ' + periodLabel(pk) + '</h3></div>';
  html += '<div class="table-wrap"><table><thead><tr><th>Metric</th><th>Target</th>';
  stores.forEach(function(st){ html += '<th class="num">' + storeName(st) + '</th>'; });
  html += '</tr></thead><tbody>';
  rows.forEach(function(r){
    html += '<tr><td><b>' + esc(r[0]) + '</b></td><td class="muted">' + r[2] + '</td>';
    stores.forEach(function(st){
      var m = metricsFor(pk, st);
      if (!m) { html += '<td class="num muted">—</td>'; return; }
      var s = scoreKPIs(m); var res = r[1](m, s);
      var cls = res.ok == null ? '' : (res.ok ? 'cell-ok bg-ok' : 'cell-bad bg-bad');
      html += '<td class="num ' + cls + '">' + res.v + '</td>';
    });
    html += '</tr>';
  });
  // totals row
  html += '<tr style="border-top:2px solid #1c1c1c"><td><b>KPIs Hit</b></td><td></td>';
  stores.forEach(function(st){
    var m = metricsFor(pk, st); var s = m ? scoreKPIs(m) : { points: 0 };
    html += '<td class="num"><span class="pill ' + (s.points>=2?'ok':'bad') + '">' + s.points + ' / 4</span></td>';
  });
  html += '</tr></tbody></table></div></div>';
  sec.innerHTML = html;
};

/* ---- Annual Summary ---- */
RENDER.annual = function(sec){
  var stores = managerStores();
  var year = 2026;
  var html = '<div class="card"><div class="section-title"><h3>Annual Summary — FY' + year + '</h3></div>';
  html += '<div class="table-wrap"><table><thead><tr><th>Store</th>';
  periodsInYear(year).forEach(function(p){ html += '<th class="num">Q' + p.q + '</th>'; });
  html += '<th class="num">YTD KPIs</th><th class="num">YTD Bonus Pool</th></tr></thead><tbody>';
  stores.forEach(function(st){
    var ytdPts = 0, ytdPool = 0;
    html += '<tr><td><b>' + storeName(st) + '</b></td>';
    periodsInYear(year).forEach(function(p){
      var m = metricsFor(p.key, st); var s = m ? scoreKPIs(m) : null;
      if (s) { ytdPts += s.points; ytdPool += techBonusPool(p.key, st).pool; }
      html += '<td class="num">' + (s ? '<span class="pill ' + (s.points>=2?'ok':'bad') + '">' + s.points + '/4</span>' : '<span class="muted">—</span>') + '</td>';
    });
    html += '<td class="num"><b>' + ytdPts + ' / 16</b></td><td class="num">' + money(ytdPool) + '</td></tr>';
  });
  html += '</tbody></table></div></div>';

  // manager YTD strip — admins see both; a manager sees only their own store
  html += '<div class="spacer"></div><div class="grid g2">';
  Object.keys(CFG.MANAGER_BONUS).filter(function(email){
    return SESSION.role === 'admin' || SESSION.store === CFG.MANAGER_BONUS[email].store;
  }).forEach(function(email){
    var y = managerYtd(email); var nm = CFG.MANAGER_BONUS[email].name;
    html += '<div class="card"><div class="section-title"><h3>' + esc(nm) + ' — KPI Bonus YTD</h3></div>' +
      progress('Earned ' + money(y.earned), 'Cap ' + money(y.cap), y.cap ? (y.earned/y.cap*100) : 0) +
      '<div class="grid g3" style="margin-top:12px">' +
        miniStat('Earned', money(y.earned), true) + miniStat('Paid', money(y.paid), true) + miniStat('Remaining', money(y.remaining), true) +
      '</div></div>';
  });
  html += '</div>';
  sec.innerHTML = html;
};

function progress(leftLabel, rightLabel, p){
  p = Math.max(0, Math.min(100, p || 0));
  return '<div class="progress-label"><span>' + esc(leftLabel) + '</span><span>' + esc(rightLabel) + '</span></div>' +
    '<div class="bar ' + (p>=100?'ok':'') + '"><i style="width:' + p + '%"></i></div>';
}

/* ---- Quarterly Results (data entry) ---- */
RENDER.quarterly = function(sec){
  var pk = CURRENT_PERIOD;
  var stores = managerStores();
  var fields = [
    ['hrs', 'Hours Reported'], ['billed', 'Hours Billed'], ['comeback', 'Comeback %'],
    ['svcGm', 'Service GM %'], ['partsGm', 'Parts GM %'], ['svcRev', 'Service Revenue $'], ['endingWip', 'Ending WIP $'],
    // growth income (parts + service, current vs prior-year same quarter)
    ['curParts', 'Parts income — this qtr $'], ['curSvc', 'Service income — this qtr $'],
    ['priorParts', 'Parts income — prior-yr qtr $'], ['priorSvc', 'Service income — prior-yr qtr $']
  ];
  var html = '<div class="card"><div class="section-title"><h3>Enter results — ' + periodLabel(pk) + '</h3>' +
    '<button class="btn btn-primary btn-sm" id="qSave">Save quarter</button></div>';
  html += '<div class="grid ' + (stores.length>1?'g2':'') + '">';
  stores.forEach(function(st){
    var m = metricsFor(pk, st) || {};
    html += '<div class="card" style="background:#fafafa"><h3 style="margin-bottom:12px">' + storeName(st) + '</h3><div class="form-grid">';
    fields.forEach(function(f){
      html += '<label class="fld"><span>' + esc(f[1]) + '</span>' +
        '<input type="number" step="any" data-store="' + st + '" data-field="' + f[0] + '" value="' + (m[f[0]] != null ? m[f[0]] : '') + '"></label>';
    });
    html += '</div>';
    // live preview
    html += '<div id="prev-' + st + '"></div></div>';
  });
  html += '</div></div>';
  sec.innerHTML = html;

  function readStore(st){
    var o = {};
    sec.querySelectorAll('input[data-store="' + st + '"]').forEach(function(inp){ o[inp.getAttribute('data-field')] = num(inp.value); });
    return o;
  }
  function refreshPrev(st){
    var data = readStore(st);
    var s = scoreKPIs(data);
    var g = growthFromData(st, data);
    g.accrual = g.gate ? (g.rate / 100) * (s.points >= 2 ? STATE.kpi.config.cap : 0) : 0;
    var gateTxt = g.combG == null ? 'enter income' :
      (g.gate ? '<span class="cell-ok">PASS ' + g.combG.toFixed(1) + '%</span>' : '<span class="cell-bad">' + g.combG.toFixed(1) + '% (need ' + g.threshold + '%)</span>');
    $('#prev-' + st, sec).innerHTML = '<div class="spacer"></div>' + kpiChips(s) +
      '<div style="margin-top:8px" class="muted">Score: <b>' + s.points + '/4</b> · Eff ' + pct(s.eff) + ' · WIP ' + pct(s.wipPct) + '</div>' +
      '<div style="margin-top:6px" class="muted">Growth: Parts ' + (g.partsG==null?'—':g.partsG.toFixed(1)+'%') + ' · Service ' + (g.svcG==null?'—':g.svcG.toFixed(1)+'%') +
      ' · Combined ' + gateTxt + (g.gate ? ' → accrue ' + money(g.accrual) : '') + '</div>';
  }
  stores.forEach(refreshPrev);
  sec.querySelectorAll('input').forEach(function(inp){
    inp.addEventListener('input', function(){ refreshPrev(inp.getAttribute('data-store')); });
  });
  $('#qSave', sec).addEventListener('click', function(){
    if (!STATE.kpi.quarters[pk]) STATE.kpi.quarters[pk] = {};
    stores.forEach(function(st){ STATE.kpi.quarters[pk][st] = readStore(st); });
    saveKPI();
  });
};

/* ---- Staff Roster ---- */
function staffRec(name){ return STATE.staff.filter(function(x){ return x.name === name; })[0]; }

RENDER.staff = function(sec){
  var admin = AUTH.isAdmin(SESSION);
  var roster = managerFilterStaff(STATE.staff);
  var activeRoster = roster.filter(function(s){ return s.active !== false; });
  var totalAnnual = activeRoster.reduce(function(a, s){ return a + annualIncome(s); }, 0);

  var html = '<div class="card"><div class="section-title"><h3>Staff Roster</h3>' +
    '<div class="flex gap">' +
      (admin ? '<button class="btn btn-ghost btn-sm" id="addStaff">+ Add</button>' : '') +
      (admin ? '<button class="btn btn-ghost btn-sm" id="savePins">Save PINs</button>' : '') +
      (admin ? '<button class="btn btn-primary btn-sm" id="saveRoster">Save roster</button>' : '') +
    '</div></div>';

  // payroll summary strip (approx annual wages)
  html += '<div class="grid g3" style="margin-bottom:16px">' +
    stat('Active employees', activeRoster.length, roster.length + ' on roster', 'blue') +
    stat('Approx. annual wages', money(totalAnnual), 'active staff, 40h × 52wk basis', 'ok') +
    stat('Avg per employee', money(activeRoster.length ? totalAnnual / activeRoster.length : 0), '', '') +
    '</div>';

  html += '<div class="table-wrap"><table><thead><tr>' +
    '<th>Name</th><th>Emp #</th><th>Store</th><th>Role</th><th class="num">FTE</th>' +
    '<th class="num">Wage</th><th class="num">Approx. Annual</th><th class="num">Vac (wks)</th>' +
    (admin ? '<th class="num">PIN</th><th>Actions</th>' : '') + '</tr></thead><tbody>';
  roster.forEach(function(s){
    var inactive = s.active === false;
    html += '<tr' + (inactive ? ' style="opacity:.5"' : '') + '><td><b>' + esc(s.name) + '</b>' +
        (inactive ? ' <span class="pill muted">inactive</span>' : '') + '</td>' +
      '<td class="mono">' + (s.empNum ? esc(s.empNum) : (s.email ? '<span class="muted" title="' + esc(s.email) + '">email</span>' : '<span class="muted">—</span>')) + '</td>' +
      '<td>' + storeName(s.store) + '</td>' +
      '<td><span class="pill muted">' + esc(s.roleType) + '</span></td>' +
      '<td class="num">' + s.fte + '</td>' +
      '<td class="num">' + (s.payRate ? wageLabel(s) : '<span class="muted">—</span>') + '</td>' +
      '<td class="num"><b>' + (s.payRate ? money(annualIncome(s)) : '—') + '</b></td>' +
      '<td class="num">' + (s.vacationWeeks || 0) + '</td>';
    if (admin) {
      html += '<td class="num"><input type="text" maxlength="4" inputmode="numeric" style="width:72px;text-align:center" ' +
        'class="mono" data-pin="' + esc(s.name) + '" value="' + esc(s.pin) + '" placeholder="––––"></td>';
      html += '<td><div class="flex gap">' +
        '<button class="btn btn-ghost btn-sm" data-edit="' + esc(s.name) + '">Edit</button>' +
        '<button class="btn btn-primary btn-sm" data-raise="' + esc(s.name) + '">Raise</button>' +
        '<button class="btn btn-ghost btn-sm" data-hist="' + esc(s.name) + '">History</button>' +
      '</div></td>';
    }
    html += '</tr>';
  });
  html += '</tbody></table></div>';
  html += '<div class="muted" style="margin-top:10px">Approx. annual = hourly rate × 40h × 52wk × FTE (salary shown as-is). Vacation is paid time off and does not reduce the estimate.</div>';
  if (!admin) html += '<div class="muted" style="margin-top:6px">Editing, PINs and pay history are admin-only.</div>';
  html += '</div>';
  sec.innerHTML = html;

  if (admin) {
    sec.querySelectorAll('input[data-pin]').forEach(function(inp){
      inp.addEventListener('change', function(){
        var rec = staffRec(inp.getAttribute('data-pin'));
        if (rec) rec.pin = inp.value.replace(/\D/g,'').slice(0,4);
        inp.value = rec ? rec.pin : '';
      });
    });
    sec.querySelectorAll('button[data-edit]').forEach(function(b){
      b.addEventListener('click', function(){ showEditStaff(b.getAttribute('data-edit')); }); });
    sec.querySelectorAll('button[data-raise]').forEach(function(b){
      b.addEventListener('click', function(){ showRaiseModal(b.getAttribute('data-raise')); }); });
    sec.querySelectorAll('button[data-hist]').forEach(function(b){
      b.addEventListener('click', function(){ showWageHistory(b.getAttribute('data-hist')); }); });
    $('#saveRoster', sec).addEventListener('click', saveStaff);
    $('#savePins', sec).addEventListener('click', savePINs);
    $('#addStaff', sec).addEventListener('click', function(){ showAddStaff(); });
  }
};

/* ---- Edit employee: wage, pay type, FTE, vacation ---- */
function showEditStaff(name){
  var s = staffRec(name); if (!s) return;
  modal('Edit — ' + name,
    '<div class="form-grid">' +
      '<label class="fld"><span>Store</span><select id="eStore">' +
        '<option value="south"' + (s.store==='south'?' selected':'') + '>South Store</option>' +
        '<option value="north"' + (s.store==='north'?' selected':'') + '>North Store</option></select></label>' +
      '<label class="fld"><span>Role</span><select id="eRole">' +
        ['tech','support','manager','admin'].map(function(r){ return '<option value="'+r+'"'+(s.roleType===r?' selected':'')+'>'+r+'</option>'; }).join('') +
        '</select></label>' +
      '<label class="fld"><span>Pay type</span><select id="ePayType">' +
        '<option value="Hourly"' + ((s.payType||'Hourly')==='Hourly'?' selected':'') + '>Hourly</option>' +
        '<option value="Salary"' + (s.payType==='Salary'?' selected':'') + '>Salary</option></select></label>' +
      '<label class="fld"><span id="eRateLbl">Hourly rate $</span><input type="number" step="0.01" id="eRate" value="' + num(s.payRate) + '"></label>' +
      '<label class="fld"><span>FTE</span><input type="number" step="0.1" id="eFte" value="' + (s.fte==null?1:s.fte) + '"></label>' +
      '<label class="fld"><span>Vacation (weeks)</span><input type="number" step="0.5" id="eVac" value="' + (s.vacationWeeks||0) + '"></label>' +
      '<label class="fld"><span>Employee # (for login &amp; efficiency join)</span><input type="text" id="eEmpNum" value="' + esc(s.empNum || '') + '" placeholder="e.g. S043"></label>' +
      '<label class="fld"><span>Email (managers/admin login)</span><input type="email" id="eEmail" value="' + esc(s.email || '') + '" placeholder="name@hydeparkequipment.ca"></label>' +
      '<label class="fld" style="grid-column:1 / -1"><span>Labour-data name (override — only if efficiency doesn’t auto-match)</span>' +
        '<input type="text" id="eEffName" value="' + esc(s.effName || '') + '" placeholder="e.g. NATHAN VERBURG"></label>' +
    '</div>' +
    '<div id="eAnnual" class="stat ok" style="margin:4px 0 14px"></div>' +
    '<div class="flex gap" style="justify-content:flex-end">' +
      '<button class="btn btn-ghost" id="eInactive">' + (s.active===false?'Reactivate':'Mark inactive') + '</button>' +
      '<button class="btn btn-primary" id="eSave">Save changes</button></div>',
  function(box){
    function preview(){
      var tmp = { payRate: num($('#eRate',box).value), payType: $('#ePayType',box).value, fte: num($('#eFte',box).value) };
      $('#eRateLbl',box).textContent = tmp.payType === 'Salary' ? 'Annual salary $' : 'Hourly rate $';
      $('#eAnnual',box).innerHTML = '<div class="label">Approx. annual income</div><div class="value" style="font-size:24px">' +
        money(annualIncome(tmp)) + '</div><div class="meta">' + ($('#eVac',box).value||0) + ' weeks vacation</div>';
    }
    ['eRate','ePayType','eFte','eVac'].forEach(function(id){ $('#'+id,box).addEventListener('input', preview); });
    preview();
    $('#eInactive',box).addEventListener('click', function(){
      s.active = (s.active === false); // toggle
      saveStaff().then(function(){ closeModal(); go('staff'); });
    });
    $('#eSave',box).addEventListener('click', function(){
      s.store = $('#eStore',box).value; s.division = storeDivision(s.store);
      s.roleType = $('#eRole',box).value;
      s.payType = $('#ePayType',box).value;
      s.payRate = num($('#eRate',box).value);
      s.fte = num($('#eFte',box).value);
      s.vacationWeeks = num($('#eVac',box).value);
      s.empNum = $('#eEmpNum',box).value.trim();
      s.email = $('#eEmail',box).value.trim();
      s.effName = $('#eEffName',box).value.trim();
      rebuildEfficiencyIndex();   // re-link efficiency with the new empNum / override
      saveStaff().then(function(){ closeModal(); go('staff'); });
    });
  });
}

/* ---- Add a new employee ---- */
function showAddStaff(){
  modal('Add Employee',
    '<div class="form-grid">' +
      '<label class="fld" style="grid-column:1 / -1"><span>Full name *</span><input type="text" id="aName" placeholder="e.g. Jane Smith" autofocus></label>' +
      '<label class="fld"><span>Store</span><select id="aStore"><option value="south">South Store</option><option value="north">North Store</option></select></label>' +
      '<label class="fld"><span>Role</span><select id="aRole">' +
        ['tech','support','manager','admin'].map(function(r){ return '<option value="'+r+'">'+r+'</option>'; }).join('') + '</select></label>' +
      '<label class="fld"><span>Pay type</span><select id="aPayType"><option value="Hourly">Hourly</option><option value="Salary">Salary</option></select></label>' +
      '<label class="fld"><span id="aRateLbl">Hourly rate $</span><input type="number" step="0.01" id="aRate" value="0"></label>' +
      '<label class="fld"><span>FTE</span><input type="number" step="0.1" id="aFte" value="1"></label>' +
      '<label class="fld"><span>Vacation (weeks)</span><input type="number" step="0.5" id="aVac" value="0"></label>' +
      '<label class="fld"><span>Employee # (tech/support login)</span><input type="text" id="aEmpNum" placeholder="e.g. S043"></label>' +
      '<label class="fld"><span>Email (manager/admin login)</span><input type="email" id="aEmail" placeholder="name@hydeparkequipment.ca"></label>' +
      '<label class="fld"><span>PIN / passcode (4 digits)</span><input type="text" maxlength="4" inputmode="numeric" id="aPin" placeholder="––––"></label>' +
    '</div>' +
    '<div id="aAnnual" class="stat ok" style="margin:4px 0 14px"></div>' +
    '<div class="msg" id="aMsg" style="display:none"></div>' +
    '<div style="text-align:right"><button class="btn btn-primary" id="aSave">Add employee</button></div>',
  function(box){
    function preview(){
      var tmp = { payRate: num($('#aRate',box).value), payType: $('#aPayType',box).value, fte: num($('#aFte',box).value) };
      $('#aRateLbl',box).textContent = tmp.payType === 'Salary' ? 'Annual salary $' : 'Hourly rate $';
      $('#aAnnual',box).innerHTML = '<div class="label">Approx. annual income</div><div class="value" style="font-size:24px">' +
        money(annualIncome(tmp)) + '</div>';
    }
    ['aRate','aPayType','aFte'].forEach(function(id){ $('#'+id,box).addEventListener('input', preview); });
    preview();
    $('#aSave',box).addEventListener('click', function(){
      var name = $('#aName',box).value.trim();
      var msg = $('#aMsg',box);
      if (!name){ msg.style.display='block'; msg.className='msg err'; msg.textContent='Enter a name.'; return; }
      if (staffRec(name)){ msg.style.display='block'; msg.className='msg err'; msg.textContent='An employee named “'+name+'” already exists.'; return; }
      var store = $('#aStore',box).value;
      var rec = {
        name: name, store: store, division: storeDivision(store),
        roleType: $('#aRole',box).value, fte: num($('#aFte',box).value),
        pin: $('#aPin',box).value.replace(/\D/g,'').slice(0,4),
        empNum: $('#aEmpNum',box).value.trim(), email: $('#aEmail',box).value.trim(), effName: '',
        payRate: num($('#aRate',box).value), payType: $('#aPayType',box).value,
        vacationWeeks: num($('#aVac',box).value), payHistory: [], active: true, queue: ''
      };
      STATE.staff.push(rec);
      rebuildEfficiencyIndex();
      closeModal();
      go('staff');          // render the new row immediately
      saveStaff();          // persist in the background (toasts result)
    });
  });
}

/* ---- Raise: annual / performance increase with live % ---- */
function showRaiseModal(name){
  var s = staffRec(name); if (!s) return;
  var isSalary = (s.payType === 'Salary');
  var unit = isSalary ? '/yr' : '/hr';
  var old = num(s.payRate);
  modal('Raise — ' + name,
    '<div class="muted" style="margin-bottom:12px">Current: <b>' + (isSalary ? money(old) : money2(old)) + unit +
      '</b> · approx. annual <b>' + money(annualIncome(s)) + '</b></div>' +
    '<div class="form-grid">' +
      '<label class="fld"><span>Increase %</span><input type="number" step="0.1" id="rPct" placeholder="e.g. 3"></label>' +
      '<label class="fld"><span>New rate ' + unit + '</span><input type="number" step="0.01" id="rNew" value="' + old + '"></label>' +
      '<label class="fld"><span>Type</span><select id="rType">' +
        '<option>Annual increase</option><option>Performance increase</option><option>Promotion</option><option>Market adjustment</option><option>Other</option>' +
        '</select></label>' +
      '<label class="fld"><span>Effective date</span><input type="text" id="rDate" value="' + todayStr() + '"></label>' +
    '</div>' +
    '<label class="fld"><span>Note (optional)</span><input type="text" id="rNote" placeholder="reason / details"></label>' +
    '<div id="rOut" class="stat" style="margin:4px 0 14px"></div>' +
    '<div style="text-align:right"><button class="btn btn-primary" id="rSave">Apply raise</button></div>',
  function(box){
    var lastEdited = 'pct';
    function recompute(src){
      if (src) lastEdited = src;
      var newRate;
      if (lastEdited === 'pct') {
        var pct = num($('#rPct',box).value);
        newRate = old * (1 + pct/100);
        $('#rNew',box).value = (Math.round(newRate*100)/100);
      } else {
        newRate = num($('#rNew',box).value);
        var p = old > 0 ? (newRate - old)/old*100 : 0;
        $('#rPct',box).value = (Math.round(p*10)/10);
      }
      var pctChange = old > 0 ? (newRate - old)/old*100 : 0;
      var tmpNew = { payRate: newRate, payType: s.payType, fte: s.fte };
      var deltaAnnual = annualIncome(tmpNew) - annualIncome(s);
      var good = pctChange >= 0;
      $('#rOut',box).className = 'stat ' + (good ? 'ok' : 'bad');
      $('#rOut',box).innerHTML =
        '<div class="label">New ' + (isSalary?'salary':'rate') + ' & change</div>' +
        '<div class="value" style="font-size:24px">' + (isSalary?money(newRate):money2(newRate)) + unit +
          ' <span style="font-size:16px;color:' + (good?'var(--ok)':'var(--bad)') + '">(' + (good?'+':'') + pctChange.toFixed(1) + '%)</span></div>' +
        '<div class="meta">New approx. annual ' + money(annualIncome(tmpNew)) +
          ' · ' + (deltaAnnual>=0?'+':'') + money(deltaAnnual) + '/yr</div>';
    }
    $('#rPct',box).addEventListener('input', function(){ recompute('pct'); });
    $('#rNew',box).addEventListener('input', function(){ recompute('new'); });
    recompute();
    $('#rSave',box).addEventListener('click', function(){
      var newRate = num($('#rNew',box).value);
      var pctChange = old > 0 ? (newRate - old)/old*100 : 0;
      s.payHistory = s.payHistory || [];
      s.payHistory.push({ date: $('#rDate',box).value, oldRate: old, newRate: newRate,
        pctChange: pctChange.toFixed(1), note: ($('#rType',box).value + ($('#rNote',box).value ? ' — ' + $('#rNote',box).value : '')) });
      s.payRate = newRate;
      saveStaff().then(function(){ closeModal(); go('staff'); });
    });
  });
}

/* ---- Wage history (rate changes) + bonus payments ---- */
function showWageHistory(name){
  var s = staffRec(name);
  var ph = (s && s.payHistory) || [];
  var rateRows = ph.length ? ph.slice().reverse().map(function(h){
    var pc = h.pctChange != null && h.pctChange !== '' ? h.pctChange + '%' : '';
    return '<tr><td>' + esc(h.date) + '</td><td class="num">' + (h.oldRate?money2(h.oldRate):'—') + '</td>' +
      '<td class="num">' + money2(h.newRate) + '</td><td class="num ' + (parseFloat(h.pctChange)>=0?'cell-ok':'cell-bad') + '">' + pc + '</td>' +
      '<td>' + esc(h.note||'') + '</td></tr>';
  }).join('') : '<tr><td colspan="5" class="muted">No wage changes recorded.</td></tr>';

  var pays = STATE.kpi.payments.filter(function(p){ return p.employee === name; });
  var payRows = pays.length ? pays.slice().reverse().map(function(p){
    return '<tr><td>' + esc(p.date) + '</td><td>' + esc(periodLabel(p.period)) + '</td><td class="num">' + money2(p.amount) + '</td><td>' + esc(p.note||'') + '</td></tr>';
  }).join('') : '<tr><td colspan="4" class="muted">No bonus payments recorded.</td></tr>';

  modal('History — ' + name,
    '<h3 style="font-size:16px;margin-bottom:8px">Wage changes</h3>' +
    '<div class="table-wrap"><table><thead><tr><th>Date</th><th class="num">From</th><th class="num">To</th><th class="num">%</th><th>Reason</th></tr></thead><tbody>' +
    rateRows + '</tbody></table></div>' +
    '<h3 style="font-size:16px;margin:18px 0 8px">Bonus payments</h3>' +
    '<div class="table-wrap"><table><thead><tr><th>Date</th><th>Period</th><th class="num">Amount</th><th>Note</th></tr></thead><tbody>' +
    payRows + '</tbody></table></div>');
}

/* ---- Payroll Summary ---- */
RENDER.payroll = function(sec){
  var pk = CURRENT_PERIOD;
  var stores = managerStores();
  var html = '<div class="card"><div class="section-title"><h3>Payroll Summary — ' + periodLabel(pk) + '</h3></div>';
  stores.forEach(function(st){
    var pool = techBonusPool(pk, st);
    html += '<div class="spacer"></div><div class="flex between center"><h3 style="font-size:17px">' + storeName(st) + '</h3>' +
      '<span class="pill ' + (pool.eligible?'ok':'bad') + '">' + (pool.eligible ? 'Eligible · ' + pool.points + '/4' : 'Not eligible (<2/4)') + '</span></div>';
    html += '<div class="grid g3" style="margin:10px 0">' +
      miniStat('Pool', money(pool.pool), pool.eligible) +
      miniStat('Tech share ' + STATE.kpi.config.techShare + '%', money(pool.techPool), true) +
      miniStat('Support share ' + (100-STATE.kpi.config.techShare) + '%', money(pool.supportPool), true) + '</div>';
    html += '<div class="table-wrap"><table><thead><tr><th>Employee</th><th>Group</th><th class="num">FTE</th><th class="num">Payout</th></tr></thead><tbody>';
    pool.lines.forEach(function(l){
      html += '<tr><td>' + esc(l.name) + '</td><td>' + l.group + '</td><td class="num">' + l.fte + '</td><td class="num"><b>' + money2(l.amount) + '</b></td></tr>';
    });
    html += '</tbody></table></div>';
  });
  html += '</div>';

  // ---- Manual / discretionary bonuses for all staff (this period) ----
  var roster = managerFilterStaff(STATE.staff).filter(function(s){ return s.active !== false; });
  var mbAll = (STATE.kpi.manualBonuses && STATE.kpi.manualBonuses[pk]) || {};
  var mbTotal = roster.reduce(function(a,s){ return a + num((mbAll[s.name] || {}).amount); }, 0);
  html += '<div class="spacer"></div><div class="card"><div class="section-title"><h3>Manual / Discretionary Bonuses — ' + periodLabel(pk) + '</h3>' +
    '<button class="btn btn-primary btn-sm" id="mbSave">Save bonuses</button></div>' +
    '<div class="muted" style="margin-bottom:10px">Enter a manual bonus for any employee this period (one-off, year-end, etc.). Total: <b>' + money2(mbTotal) + '</b></div>' +
    '<div class="table-wrap"><table><thead><tr><th>Employee</th><th>Store</th><th>Role</th><th class="num">Bonus $</th><th>Note</th></tr></thead><tbody>';
  roster.forEach(function(s){
    var rec = mbAll[s.name] || { amount: '', note: '' };
    html += '<tr><td><b>' + esc(s.name) + '</b></td><td>' + storeName(s.store) + '</td><td><span class="pill muted">' + esc(s.roleType) + '</span></td>' +
      '<td class="num"><input type="number" step="0.01" data-mbonus="' + esc(s.name) + '" value="' + (rec.amount !== '' && rec.amount != null ? num(rec.amount) : '') + '" placeholder="0" style="width:96px;text-align:right"></td>' +
      '<td><input type="text" data-mbnote="' + esc(s.name) + '" value="' + esc(rec.note || '') + '" placeholder="reason…" style="min-width:160px"></td></tr>';
  });
  html += '</tbody></table></div></div>';
  sec.innerHTML = html;

  function mbRec(name){
    STATE.kpi.manualBonuses[pk] = STATE.kpi.manualBonuses[pk] || {};
    STATE.kpi.manualBonuses[pk][name] = STATE.kpi.manualBonuses[pk][name] || { amount: 0, note: '' };
    return STATE.kpi.manualBonuses[pk][name];
  }
  sec.querySelectorAll('input[data-mbonus]').forEach(function(inp){
    inp.addEventListener('change', function(){ mbRec(inp.getAttribute('data-mbonus')).amount = num(inp.value); });
  });
  sec.querySelectorAll('input[data-mbnote]').forEach(function(t){
    t.addEventListener('change', function(){ mbRec(t.getAttribute('data-mbnote')).note = t.value; });
  });
  $('#mbSave', sec).addEventListener('click', function(){ saveKPI().then(function(){ if (CURRENT_PAGE) go('payroll'); }); });
};

/* ---- Tech Efficiency ---- */
RENDER.efficiency = function(sec){
  var c = STATE.kpi.config;
  var target = c.techEff;                 // define target BEFORE using it (fix #2)
  var stores = managerStores();
  if (!hasEfficiency()) {
    sec.innerHTML = '<div class="empty">No Labour efficiency data uploaded yet. Upload it on the <b>Configuration</b> page.</div>';
    return;
  }
  var months = effMonths();
  var exMap = excludedNBMap();
  var html = '';
  stores.forEach(function(st){
    var div = storeDivision(st);
    var techs = effTechsForDivision(div);   // filter techs by store division (fix #3)
    html += '<div class="card"><div class="section-title"><h3>' + storeName(st) + ' — Tech Efficiency</h3>' +
      '<span class="muted">target ' + target + '%</span></div>';
    if (!techs.length) { html += '<div class="empty">No tech rows for this store.</div></div>'; return; }
    html += '<div class="table-wrap"><table><thead><tr><th>Technician</th>';
    months.forEach(function(m){ html += '<th class="num">' + esc(m) + '</th>'; });
    html += '<th class="num">Avg</th><th class="num">Non-Bill hrs</th></tr></thead><tbody>';
    techs.forEach(function(name){
      html += '<tr><td><b>' + esc(name) + '</b></td>';
      var sumB = 0, sumR = 0, sumNB = 0;
      months.forEach(function(m){
        var cell = STATE.efficiency.byTechMonth[name] && STATE.efficiency.byTechMonth[name][m];
        if (cell) { sumNB += cell.nonBillable || 0; }
        var rep = cell ? cell.reported - (exMap[name + '|' + m] || 0) : 0;   // exclusion-adjusted
        if (cell && rep > 0) {
          var e = cell.billed / rep * 100; sumB += cell.billed; sumR += rep;
          html += '<td class="num ' + (e >= target ? 'eff-good' : 'eff-bad') + '">' + e.toFixed(0) + '%</td>';
        } else { html += '<td class="num muted">—</td>'; }
      });
      var avg = sumR > 0 ? (sumB / sumR * 100) : null;
      html += '<td class="num ' + (avg != null && avg >= target ? 'eff-good' : 'eff-bad') + '"><b>' + (avg != null ? avg.toFixed(0) + '%' : '—') + '</b></td>' +
        '<td class="num">' + (sumNB > 0 ? sumNB.toFixed(1) : '—') + '</td></tr>';
    });
    html += '</tbody></table></div></div><div class="spacer"></div>';
  });
  // flagged WOs + non-billable review/exclude list
  html += renderFlaggedWOs(stores);
  html += renderNonBillableReview(stores);
  sec.innerHTML = html;
  wireNonBillableExcludes(sec);
};

/* Non-billable review list — managers/admin can exclude entries that were
 * out of the tech's control (e.g. management-directed shop repairs), removing
 * them from the efficiency denominator. */
function renderNonBillableReview(stores){
  var divs = stores.map(storeDivision);
  var rows = (STATE.efficiency.nonBillable || []).filter(function(r){ return divs.indexOf(r.division) !== -1; });
  if (!rows.length) return '';
  var excl = STATE.kpi.effExclusions || {};
  var ord = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  rows = rows.slice().sort(function(a,b){ return (a.display||'').localeCompare(b.display||'') || (ord.indexOf(a.month)-ord.indexOf(b.month)) || (a.category||'').localeCompare(b.category||''); });
  var totalNB = rows.reduce(function(a,r){ return a + r.hours; }, 0);
  var exclNB = rows.reduce(function(a,r){ return a + nbExcludedHours(r); }, 0);
  var haveDetail = (STATE.nbDetail || []).length > 0;
  var html = '<div class="card"><div class="section-title"><h3>Non-Billable Hours — Review &amp; Exclude</h3>' +
    '<button class="btn btn-primary btn-sm" id="nbSave">Save exclusions</button></div>' +
    '<div class="muted" style="margin-bottom:10px">Non-billable hours count toward efficiency. Enter the <b>hours to exclude</b> for any entry outside the tech’s control (removed from the calc)' +
    (haveDetail ? ', or use <b>Pick</b> to choose individual time punches' : '') + '. ' +
    rows.length + ' categories · ' + totalNB.toFixed(1) + ' hrs total · ' + exclNB.toFixed(1) + ' hrs excluded.</div>' +
    '<div class="table-wrap"><table><thead><tr><th>Technician</th><th>Month</th><th>Category</th><th class="num">Entries</th><th class="num">Hours</th><th class="num">Exclude hrs</th>' +
    (haveDetail ? '<th>Punches</th>' : '') + '<th>Note</th></tr></thead><tbody>';
  rows.forEach(function(r){
    var ex = excl[r.key] || {};
    var exHrs = nbExcludedHours(r);
    html += '<tr' + (exHrs > 0 ? ' class="bg-bad"' : '') + '><td>' + esc(r.display) + '</td><td>' + esc(r.month) + '</td>' +
      '<td>' + esc(r.category) + '</td><td class="num">' + (r.count || 1) + '</td><td class="num">' + r.hours.toFixed(1) + '</td>' +
      '<td class="num"><input type="number" step="0.1" min="0" max="' + r.hours + '" data-nbhrs="' + esc(r.key) + '" value="' + (exHrs ? exHrs.toFixed(1) : '') + '" placeholder="0" style="width:74px;text-align:right"></td>' +
      (haveDetail ? '<td>' + ((STATE.nbDetail || []).some(function(d){ return d.catKey === r.key; }) ? '<button class="btn btn-ghost btn-sm" data-nbpick="' + esc(r.key) + '">Pick…</button>' : '') + '</td>' : '') +
      '<td><input type="text" data-nbnote="' + esc(r.key) + '" value="' + esc(ex.note || '') + '" placeholder="reason…" style="min-width:140px"></td></tr>';
  });
  html += '</tbody></table></div></div>';
  return html;
}
function nbRec(k){ STATE.kpi.effExclusions[k] = STATE.kpi.effExclusions[k] || { excludeHours: 0, note: '' }; return STATE.kpi.effExclusions[k]; }
function wireNonBillableExcludes(sec){
  sec.querySelectorAll('input[data-nbhrs]').forEach(function(inp){
    inp.addEventListener('change', function(){
      var k = inp.getAttribute('data-nbhrs'); var rec = nbRec(k);
      rec.excludeHours = Math.max(0, Math.min(num(inp.value), num(inp.max)));
      if ('excluded' in rec) delete rec.excluded;
      inp.value = rec.excludeHours ? rec.excludeHours.toFixed(1) : '';
    });
  });
  sec.querySelectorAll('input[data-nbnote]').forEach(function(t){
    t.addEventListener('change', function(){ nbRec(t.getAttribute('data-nbnote')).note = t.value; });
  });
  sec.querySelectorAll('button[data-nbpick]').forEach(function(b){
    b.addEventListener('click', function(){ showNonBillablePicker(b.getAttribute('data-nbpick')); });
  });
  var sv = $('#nbSave', sec);
  if (sv) sv.addEventListener('click', function(){ saveKPI().then(function(){ if (CURRENT_PAGE) go(CURRENT_PAGE); }); });
}

/* Pick individual time punches within a non-billable category; the sum of the
 * chosen punches becomes that category's "hours to exclude". */
function showNonBillablePicker(catKey){
  var punches = (STATE.nbDetail || []).filter(function(d){ return d.catKey === catKey; });
  if (!punches.length) { toast('Punch detail is only available right after an upload.', 'bad'); return; }
  var parts = catKey.split('|');
  var rec = nbRec(catKey);
  var alreadyExcl = (rec.excludeHours != null) ? num(rec.excludeHours) : 0;
  var body = '<div class="muted" style="margin-bottom:10px">' + esc(parts[0]) + ' · ' + esc(parts[1]) + ' · ' + esc(parts[2]) +
      ' — tick the punches that were outside the tech’s control. <span id="npSum">0.0</span> hrs selected.</div>' +
    '<div class="table-wrap" style="max-height:46vh;overflow:auto"><table><thead><tr><th>Exclude</th><th class="num">Hours</th><th>WO#</th><th>Comment</th></tr></thead><tbody>';
  punches.forEach(function(p, i){
    body += '<tr><td style="text-align:center"><input type="checkbox" data-np="' + i + '" data-h="' + p.hours + '"></td>' +
      '<td class="num">' + p.hours.toFixed(2) + '</td><td class="mono">' + esc(p.docNum || '') + '</td><td>' + esc(p.comment || '') + '</td></tr>';
  });
  body += '</tbody></table></div>' +
    '<div class="flex between center" style="margin-top:12px"><div class="muted">Currently excluded: ' + alreadyExcl.toFixed(1) + ' hrs</div>' +
    '<button class="btn btn-primary" id="npApply">Apply selected hours</button></div>';
  modal('Pick time punches', body, function(box){
    function sum(){ var t = 0; box.querySelectorAll('input[data-np]:checked').forEach(function(cb){ t += num(cb.getAttribute('data-h')); }); $('#npSum', box).textContent = t.toFixed(1); return t; }
    box.querySelectorAll('input[data-np]').forEach(function(cb){ cb.addEventListener('change', sum); });
    $('#npApply', box).addEventListener('click', function(){
      var rec2 = nbRec(catKey); rec2.excludeHours = Math.round(sum() * 100) / 100; if ('excluded' in rec2) delete rec2.excluded;
      saveKPI().then(function(){ closeModal(); if (CURRENT_PAGE) go(CURRENT_PAGE); });
    });
  });
}

// Merge flagged entries that share the same tech+month+WO# (handles data
// persisted as individual punches before WO-consolidation was added).
function consolidateFlagged(arr){
  var m = {};
  (arr || []).forEach(function(r){
    var nm = r.display || r.name;
    var k = nm + '|' + r.month + '|' + r.docNum;
    var w = m[k] || (m[k] = { display: nm, division: r.division, month: r.month, docNum: r.docNum, reported: 0, billed: 0, punches: 0 });
    w.reported += r.reported; w.billed += r.billed; w.punches += (r.punches || 1);
  });
  return Object.keys(m).map(function(k){ var w = m[k]; w.eff = w.reported > 0 ? w.billed / w.reported * 100 : 0; return w; });
}

function renderFlaggedWOs(stores){
  var c = STATE.kpi.config;
  // consolidate by WO, then keep only under-target (also screens out over-100%)
  var flagged = consolidateFlagged(STATE.efficiency.flagged).filter(function(r){ return r.eff < c.effLowFlag; });
  if (stores) flagged = flagged.filter(function(r){ return stores.indexOf(divisionStore(r.division)) !== -1; });
  if (!flagged.length) return '';
  // group by technician; within each tech, in calendar-month sequence
  flagged.sort(function(a,b){ return (a.display||'').localeCompare(b.display||'') || (monthIndex(a.month) - monthIndex(b.month)); });
  var html = '<div class="card"><div class="section-title"><h3>Flagged Work Orders</h3>' +
    '<span class="muted">under ' + c.effLowFlag + '% · by technician</span></div>' +
    '<div class="table-wrap"><table><thead><tr><th>Technician</th><th>Month</th><th>WO#</th><th class="num">Punches</th><th class="num">Reported</th><th class="num">Billed</th><th class="num">Eff</th><th>Flag</th></tr></thead><tbody>';
  var lastName = null;
  flagged.slice(0, 300).forEach(function(r){
    var nm = r.display || r.name;
    var newGroup = nm !== lastName; lastName = nm;
    html += '<tr class="bg-bad"' + (newGroup ? ' style="border-top:2px solid #1c1c1c"' : '') + '>' +
      '<td>' + (newGroup ? '<b>' + esc(nm) + '</b>' : '') + '</td><td>' + esc(r.month) + '</td><td class="mono">' + esc(r.docNum) + '</td>' +
      '<td class="num">' + (r.punches || 1) + '</td>' +
      '<td class="num">' + r.reported.toFixed(1) + '</td><td class="num">' + r.billed.toFixed(1) + '</td>' +
      '<td class="num"><b>' + Math.floor(r.eff) + '%</b></td>' +
      '<td><span class="pill bad">Under ' + c.effLowFlag + '%</span></td></tr>';
  });
  html += '</tbody></table></div></div>';
  return html;
}

/* ---- Payment History ---- */
RENDER.payments = function(sec){
  var html = '<div class="card"><div class="section-title"><h3>Payment History</h3>' +
    '<button class="btn btn-primary btn-sm" id="addPay">+ Record payment</button></div>';
  html += '<div class="table-wrap"><table><thead><tr><th>Date</th><th>Employee</th><th>Period</th><th class="num">Amount</th><th>Note</th></tr></thead><tbody>';
  if (!STATE.kpi.payments.length) html += '<tr><td colspan="5" class="muted">No payments recorded yet.</td></tr>';
  STATE.kpi.payments.slice().reverse().forEach(function(p){
    html += '<tr><td>' + esc(p.date) + '</td><td><b>' + esc(p.employee) + '</b></td><td>' + esc(periodLabel(p.period)) + '</td><td class="num">' + money2(p.amount) + '</td><td>' + esc(p.note||'') + '</td></tr>';
  });
  html += '</tbody></table></div></div>';
  sec.innerHTML = html;
  $('#addPay', sec).addEventListener('click', function(){
    var opts = STATE.staff.map(function(s){ return '<option>' + esc(s.name) + '</option>'; }).join('');
    var qopts = PERIODS.map(function(p){ return '<option value="' + p.key + '">' + p.label + '</option>'; }).join('');
    modal('Record payment',
      '<div class="form-grid">' +
        '<label class="fld"><span>Employee</span><select id="pEmp">' + opts + '</select></label>' +
        '<label class="fld"><span>Period</span><select id="pPer">' + qopts + '</select></label>' +
        '<label class="fld"><span>Amount $</span><input type="number" id="pAmt" step="0.01"></label>' +
        '<label class="fld"><span>Date</span><input type="text" id="pDate" value="' + todayStr() + '"></label>' +
      '</div><label class="fld"><span>Note</span><input type="text" id="pNote"></label>' +
      '<div style="text-align:right"><button class="btn btn-primary" id="pSave">Save payment</button></div>',
    function(box){
      $('#pSave', box).addEventListener('click', function(){
        STATE.kpi.payments.push({ date: $('#pDate',box).value, employee: $('#pEmp',box).value,
          period: $('#pPer',box).value, amount: num($('#pAmt',box).value), note: $('#pNote',box).value });
        saveKPI().then(function(){ closeModal(); go('payments'); });
      });
    });
  });
};

/* ---- Growth Bonus ---- */
RENDER.growth = function(sec){
  var g = STATE.kpi.growth;
  var admin = AUTH.isAdmin(SESSION);
  var viewStores = managerStores();
  var html = '';

  // ---- settings (admin) ----
  if (admin) {
    html += '<div class="card"><div class="section-title"><h3>Growth Bonus — Settings</h3>' +
      '<button class="btn btn-primary btn-sm" id="gSave">Save settings</button></div>' +
      '<div class="desc">Each store accrues ' + g.rate + '% of its tech+support bonus pool to a growth bank in any quarter its combined parts+service income grows ≥ ' + (g.threshold!=null?g.threshold:4) + '% vs the prior-year same quarter. Banks combine into one year-end pool.</div>' +
      '<div class="form-grid">' +
        '<label class="fld"><span>Accrual rate %</span><input type="number" id="gRate" value="' + g.rate + '"></label>' +
        '<label class="fld"><span>Growth threshold %</span><input type="number" id="gThresh" value="' + (g.threshold!=null?g.threshold:4) + '"></label>' +
      '</div></div><div class="spacer"></div>';
  }

  // ---- per-store accrual breakdown ----
  viewStores.forEach(function(st){
    var bank = storeBank(st);
    html += '<div class="card"><div class="section-title"><h3>' + storeName(st) + ' — Growth Bank</h3>' +
      '<span class="pill ' + (bank>0?'ok':'muted') + '">' + money(bank) + ' accrued</span></div>' +
      '<div class="table-wrap"><table><thead><tr><th>Quarter</th><th class="num">Parts growth</th><th class="num">Service growth</th><th class="num">Combined</th><th>Gate (' + (g.threshold!=null?g.threshold:4) + '%)</th><th class="num">Tech+Support pool</th><th class="num">Accrual ' + g.rate + '%</th></tr></thead><tbody>';
    periodsInYear(2026).forEach(function(p){
      var sg = storeGrowth(p.key, st);
      var fmt = function(x){ return x==null?'<span class="muted">—</span>':'<span class="' + (x>=0?'cell-ok':'cell-bad') + '">' + x.toFixed(1) + '%</span>'; };
      html += '<tr><td><b>Q' + p.q + '</b></td><td class="num">' + fmt(sg.partsG) + '</td><td class="num">' + fmt(sg.svcG) + '</td><td class="num">' + fmt(sg.combG) + '</td>' +
        '<td>' + (sg.combG==null?'<span class="muted">—</span>':(sg.gate?'<span class="pill ok">PASS</span>':'<span class="pill bad">miss</span>')) + '</td>' +
        '<td class="num">' + money(sg.pool) + '</td><td class="num"><b>' + money(sg.accrual) + '</b></td></tr>';
    });
    html += '<tr style="border-top:2px solid #1c1c1c"><td colspan="6"><b>Store growth bank (year)</b></td><td class="num"><b>' + money(bank) + '</b></td></tr>';
    html += '</tbody></table></div><div class="muted" style="margin-top:8px">Enter the parts/service income on <b>Quarterly Results</b> to drive these numbers.</div></div><div class="spacer"></div>';
  });

  // ---- year-end combined pool + allocation (admin only) ----
  if (admin) html += renderYearEndPool();
  sec.innerHTML = html;

  if (admin) {
    var gs = $('#gSave', sec);
    if (gs) gs.addEventListener('click', function(){ g.rate = num($('#gRate',sec).value); g.threshold = num($('#gThresh',sec).value); saveKPI().then(function(){ go('growth'); }); });
    wireYearEndPool(sec);
  }
};

/* ---- Year-end combined pool, split by manual ratio across all staff ---- */
function renderYearEndPool(){
  var g = STATE.kpi.growth;
  var pool = combinedGrowthPool();
  var remaining = Math.max(0, pool - num(g.paidOut));
  // eligible = active techs + support + warranty admin (company-wide)
  var elig = STATE.staff.filter(function(s){ return s.active !== false && ['tech','support','admin'].indexOf(s.roleType) !== -1; });
  // sort: techs, support, admin
  var order = { tech:0, support:1, admin:2 };
  elig.sort(function(a,b){ return (order[a.roleType]-order[b.roleType]) || a.name.localeCompare(b.name); });
  var alloc = g.allocation || {};
  var totalRatio = elig.reduce(function(a,s){ return a + num(alloc[s.name]); }, 0);
  var html = '<div class="card"><div class="section-title"><h3>Year-End Growth Pool — Allocation</h3>' +
    '<button class="btn btn-primary btn-sm" id="yeSave">Save allocation</button></div>' +
    '<div class="grid g3" style="margin-bottom:12px">' +
      stat('Combined pool', money(pool), 'South + North banks', 'ok') +
      stat('Paid out', money(g.paidOut), '', '') +
      stat('Remaining', money(remaining), '', 'blue') + '</div>' +
    '<label class="fld" style="max-width:240px"><span>Total paid out $</span><input type="number" id="yePaid" value="' + num(g.paidOut) + '"></label>' +
    '<div class="muted" style="margin:6px 0 10px">Enter a ratio (or %) per person — Jared can be weighted higher than Ed. Each share = pool × (their ratio ÷ total ratio). Total ratio: <b>' + totalRatio + '</b></div>' +
    '<div class="table-wrap"><table><thead><tr><th>Employee</th><th>Store</th><th>Role</th><th class="num">Ratio / %</th><th class="num">Share $</th></tr></thead><tbody>';
  elig.forEach(function(s){
    var ratio = num(alloc[s.name]);
    var share = totalRatio > 0 ? pool * (ratio / totalRatio) : 0;
    html += '<tr><td><b>' + esc(s.name) + '</b></td><td>' + storeName(s.store) + '</td><td><span class="pill muted">' + esc(s.roleType) + '</span></td>' +
      '<td class="num"><input type="number" step="0.1" min="0" data-ye="' + esc(s.name) + '" value="' + (ratio || '') + '" placeholder="0" style="width:84px;text-align:right"></td>' +
      '<td class="num"><b>' + money2(share) + '</b></td></tr>';
  });
  html += '</tbody></table></div></div>';
  return html;
}
function wireYearEndPool(sec){
  var g = STATE.kpi.growth;
  var pd = $('#yePaid', sec);
  if (pd) pd.addEventListener('change', function(){ g.paidOut = num(pd.value); });
  sec.querySelectorAll('input[data-ye]').forEach(function(inp){
    inp.addEventListener('change', function(){ g.allocation = g.allocation || {}; g.allocation[inp.getAttribute('data-ye')] = num(inp.value); if (CURRENT_PAGE) go('growth'); });
  });
  var sv = $('#yeSave', sec);
  if (sv) sv.addEventListener('click', function(){ if (pd) g.paidOut = num(pd.value); saveKPI().then(function(){ go('growth'); }); });
}

/* ---- Manager Bonuses ---- */
RENDER['manager-bonuses'] = function(sec){
  var pk = CURRENT_PERIOD;
  // NOTE: deliberately NO efficiency bars on this page (fix #5).
  var html = '';
  var emails = Object.keys(CFG.MANAGER_BONUS).filter(function(email){
    if (SESSION.role === 'admin') return true;
    return SESSION.store === CFG.MANAGER_BONUS[email].store; // manager sees own
  });
  emails.forEach(function(email){
    var b = managerBonusFor(email, pk);
    var y = managerYtd(email);
    var nm = CFG.MANAGER_BONUS[email].name;
    html += '<div class="card"><div class="section-title"><h3>' + esc(nm) + ' · ' + storeName(b.store) + '</h3>' +
      '<div class="flex gap center">' +
      (AUTH.isAdmin(SESSION) ? '<label class="muted" style="font-size:12px;display:flex;align-items:center;gap:6px">Annual cap $<input type="number" step="100" data-cap="' + email + '" value="' + ((STATE.kpi.managerBonus[email] && STATE.kpi.managerBonus[email].annualCap) || 0) + '" style="width:96px"></label>' : '') +
      '<button class="btn btn-ghost btn-sm" data-pdf="' + email + '">Export PDF</button>' +
      '<button class="btn ' + (b.paid?'btn-ghost':'btn-primary') + ' btn-sm" data-paid="' + email + '">' + (b.paid ? '✓ Paid' : 'Mark paid') + '</button></div></div>';

    // formula breakdown cards
    html += '<div class="grid g4">' +
      stat('Annual cap', money(b.cap), '', '') +
      stat('Quarterly max', money(b.quarterlyMax), 'cap ÷ 4', 'blue') +
      stat('Per KPI', money(b.perKpi), 'qtr ÷ 4 KPIs', 'blue') +
      stat('This quarter', money(b.payout), b.score.points + '/4 hit', b.payout>0?'ok':'bad') + '</div>';

    // KPI ticks for the period
    html += '<div class="spacer"></div>' + kpiChips(b.score);

    // progress to annual cap
    html += '<div class="spacer"></div>' + progress('YTD earned ' + money(y.earned), 'Annual cap ' + money(y.cap), y.cap?(y.earned/y.cap*100):0);

    // formula explanation
    html += '<div class="card" style="margin-top:14px;background:#fafafa"><b>How it’s calculated</b>' +
      '<div class="muted" style="margin-top:6px">Payout = KPIs hit × (annual cap ÷ 4 quarters ÷ 4 KPIs). ' +
      'At ' + money(b.cap) + '/yr that is ' + money(b.perKpi) + ' per KPI, ' + money(b.quarterlyMax) + ' max per quarter. ' +
      'This quarter ' + b.score.points + '/4 hit = <b>' + money(b.payout) + '</b>.</div></div>';
    html += '</div><div class="spacer"></div>';
  });

  // Year-end non-KPI staff bonuses note
  html += '<div class="card"><h3>Year-end staff bonuses (non-KPI)</h3>' +
    '<div class="muted" style="margin-top:6px">Discretionary year-end bonuses for non-KPI staff are recorded on the <b>Payment History</b> page and accrue from the <b>Growth Bonus</b> bank.</div></div>';
  sec.innerHTML = html;

  sec.querySelectorAll('button[data-paid]').forEach(function(btn){
    btn.addEventListener('click', function(){
      var email = btn.getAttribute('data-paid');
      var mb = STATE.kpi.managerBonus[email]; if (!mb.paid) mb.paid = {};
      mb.paid[pk] = !mb.paid[pk];
      saveKPI().then(function(){ go('manager-bonuses'); });
    });
  });
  sec.querySelectorAll('button[data-pdf]').forEach(function(btn){
    btn.addEventListener('click', function(){ exportManagerPDF(btn.getAttribute('data-pdf')); });
  });
  sec.querySelectorAll('input[data-cap]').forEach(function(inp){
    inp.addEventListener('change', function(){
      var email = inp.getAttribute('data-cap');
      STATE.kpi.managerBonus[email] = STATE.kpi.managerBonus[email] || { paid: {} };
      STATE.kpi.managerBonus[email].annualCap = num(inp.value);
      saveKPI().then(function(){ go('manager-bonuses'); });
    });
  });
};

/* ---- Scenario Testing ---- */
RENDER.scenario = function(sec){
  var c = STATE.kpi.config;
  var html = '<div class="card"><div class="section-title"><h3>Scenario — what-if calculator</h3></div>' +
    '<div class="desc">Enter hypothetical numbers to see the KPI score and the resulting bonus pool. Nothing is saved.</div>' +
    '<div class="form-grid">' +
    field('sHrs','Hours Reported',100) + field('sBilled','Hours Billed',80) + field('sCb','Comeback %',1.5) +
    field('sSvc','Service GM %',79) + field('sParts','Parts GM %',33) + field('sRev','Service Revenue $',500000) + field('sWip','Ending WIP $',8000) +
    '</div><div id="sOut"></div></div>';
  sec.innerHTML = html;
  function recompute(){
    var m = { hrs:num($('#sHrs',sec).value), billed:num($('#sBilled',sec).value), comeback:num($('#sCb',sec).value),
      svcGm:num($('#sSvc',sec).value), partsGm:num($('#sParts',sec).value), svcRev:num($('#sRev',sec).value), endingWip:num($('#sWip',sec).value) };
    var s = scoreKPIs(m);
    var pool = s.points >= 2 ? c.cap : 0;
    $('#sOut',sec).innerHTML = '<div class="spacer"></div>' + kpiChips(s) +
      '<div class="grid g3" style="margin-top:12px">' +
      miniStat('KPIs hit', s.points + '/4', s.points>=2) +
      miniStat('Efficiency', pct(s.eff), s.eff>=c.techEff) +
      miniStat('Bonus pool', money(pool), s.points>=2) + '</div>';
  }
  sec.querySelectorAll('input').forEach(function(i){ i.addEventListener('input', recompute); });
  recompute();
};
function field(id,label,val){ return '<label class="fld"><span>' + esc(label) + '</span><input type="number" step="any" id="' + id + '" value="' + val + '"></label>'; }

/* ============================ WARRANTY ADMIN KPI ============================
 * From the Warranty Work Orders upload. Per quarter (by close date): % of
 * warranty WOs closed within 14 days of last tech punch -> graduated rating.
 * Individual WOs outside the admin's control can be excluded. */
function warrantyKpiForQuarter(periodKey){
  var wa = STATE.kpi.warrantyAdmin || {};
  var excl = wa.exclusions || {};
  var all = (STATE.warrantyWO.rows || []).filter(function(w){ return w.quarter === periodKey; });
  var counted = all.filter(function(w){ return !(excl[w.doc] && excl[w.doc].excluded); });
  var total = counted.length;
  var within14 = counted.filter(function(w){ return w.lastPunchDays <= 14; }).length;
  var pctv = total > 0 ? within14 / total * 100 : null;
  var rating = pctv == null ? 0 : pctv >= 90 ? 100 : pctv >= 75 ? 70 : pctv >= 50 ? 40 : 15;
  var bonusAmt = num(wa.bonusAmount) || 5000;
  var quarterlyMax = bonusAmt / 4;
  var payout = quarterlyMax * rating / 100;
  var over30 = counted.filter(function(w){ return w.lastPunchDays > 30; });
  var sales = counted.reduce(function(a,w){ return a + num(w.total); }, 0);
  var internal = counted.reduce(function(a,w){ return w.internal ? a + num(w.total) : a; }, 0);
  return {
    all: all, counted: counted, total: total, within14: within14, pct: pctv, rating: rating,
    bonusAmt: bonusAmt, quarterlyMax: quarterlyMax, payout: payout,
    over30: over30, internalPct: sales > 0 ? internal / sales * 100 : 0, sales: sales, internal: internal,
    met: pctv != null && pctv >= 90
  };
}
function warrantyKpiAnnual(){
  var quarters = periodsInYear(2026).map(function(p){ return warrantyKpiForQuarter(p.key); });
  var base = quarters.reduce(function(a,q){ return a + q.payout; }, 0);
  var metCount = quarters.filter(function(q){ return q.met; }).length;
  var multiplier = metCount >= 3 ? 1.04 : 1;           // 4% growth multiplier if >=3/4 met
  return { quarters: quarters, base: base, metCount: metCount, multiplier: multiplier, total: base * multiplier };
}

RENDER['warranty-kpi'] = function(sec){
  var wa = STATE.kpi.warrantyAdmin || (STATE.kpi.warrantyAdmin = { bonusAmount: 5000, exclusions: {}, paid: {} });
  var admin = AUTH.isAdmin(SESSION);
  var ann = warrantyKpiAnnual();
  var nWO = (STATE.warrantyWO.rows || []).length;

  var html = '<div class="card"><div class="section-title"><h3>Warranty Admin KPI — FY2026</h3>' +
    (admin ? '<button class="btn btn-primary btn-sm" id="waSave">Save</button>' : '') + '</div>' +
    '<div class="desc">A warranty WO is closed the day the claim is submitted; performance = days from the last tech punch to close. Target: ≥ 90% closed within 14 days. ' + nWO + ' warranty WOs loaded.</div>' +
    '<div class="form-grid">' +
      '<label class="fld"><span>Annual bonus amount $</span><input type="number" id="waBonus" value="' + (num(wa.bonusAmount)||5000) + '"' + (admin?'':' disabled') + '></label>' +
    '</div></div><div class="spacer"></div>';

  if (!nWO) html += '<div class="empty">No warranty work orders uploaded yet. Upload them on the <b>Configuration</b> page (Warranty Work Orders).</div>';
  else {
    // top stats
    html += '<div class="grid g4">' +
      stat('Annual payout', money(ann.total), ann.multiplier > 1 ? '×1.04 growth applied' : 'base', 'ok') +
      stat('Quarters met', ann.metCount + ' / 4', '≥90% target', ann.metCount>=3?'ok':'bad') +
      stat('Growth multiplier', ann.multiplier > 1 ? '+4%' : 'none', '≥3/4 quarters', ann.multiplier>1?'ok':'') +
      stat('Bonus amount', money(num(wa.bonusAmount)||5000), money(((num(wa.bonusAmount)||5000)/4)) + '/qtr max', '') + '</div><div class="spacer"></div>';

    // per-quarter breakdown
    html += '<div class="card"><div class="section-title"><h3>Quarterly Performance</h3></div>' +
      '<div class="table-wrap"><table><thead><tr><th>Quarter</th><th class="num">Warranty WOs</th><th class="num">≤14 days</th><th class="num">% within 14d</th><th class="num">Rating</th><th class="num">Payout</th><th>Guardrails</th><th>' + (admin?'Status':'') + '</th></tr></thead><tbody>';
    periodsInYear(2026).forEach(function(p){
      var q = warrantyKpiForQuarter(p.key);
      var guard = [];
      if (q.over30.length) guard.push('<span class="pill bad">' + q.over30.length + ' over 30d</span>');
      if (q.internalPct > 10) guard.push('<span class="pill warn">int ' + q.internalPct.toFixed(0) + '%</span>');
      if (!guard.length && q.total) guard.push('<span class="pill ok">ok</span>');
      var paid = wa.paid && wa.paid[p.key];
      html += '<tr><td><b>Q' + p.q + '</b></td><td class="num">' + q.total + '</td><td class="num">' + q.within14 + '</td>' +
        '<td class="num">' + (q.pct==null?'—':'<span class="' + (q.pct>=90?'cell-ok':'cell-bad') + '">' + q.pct.toFixed(0) + '%</span>') + '</td>' +
        '<td class="num"><b>' + q.rating + '%</b></td><td class="num">' + money(q.payout) + '</td><td>' + guard.join(' ') + '</td>' +
        '<td>' + (q.total ? (admin ? '<button class="btn ' + (paid?'btn-ghost':'btn-primary') + ' btn-sm" data-wapaid="' + p.key + '">' + (paid?'✓ Paid':'Mark paid') + '</button>' : (paid?'Paid':'')) : '') + '</td></tr>';
    });
    html += '<tr style="border-top:2px solid #1c1c1c"><td colspan="5"><b>Annual (×' + ann.multiplier + ')</b></td><td class="num"><b>' + money(ann.total) + '</b></td><td colspan="2"></td></tr>';
    html += '</tbody></table></div><div class="muted" style="margin-top:8px">Rating: ≥90% → 100% · 75–89% → 70% · 50–74% → 40% · &lt;50% → 15%. Hard control: no WO open >30 days. Internal warranty should stay ≤10% of warranty sales.</div></div>';

    // exclusions: WOs that missed the 14-day target this period
    var pk = CURRENT_PERIOD;
    var misses = (STATE.warrantyWO.rows || []).filter(function(w){ return w.quarter === pk && w.lastPunchDays > 14; });
    html += '<div class="spacer"></div><div class="card"><div class="section-title"><h3>Exceptions — ' + periodLabel(pk) + ' (missed 14-day)</h3>' +
      (admin ? '<button class="btn btn-primary btn-sm" id="waExSave">Save exclusions</button>' : '') + '</div>' +
      '<div class="muted" style="margin-bottom:10px">Exclude WOs that exceeded 14 days for reasons outside the admin’s control (parts backorder, manufacturer delay, etc.) — they’re removed from this quarter’s calculation.</div>';
    if (!misses.length) html += '<div class="empty">No warranty WOs missed the 14-day target this quarter. 👍</div>';
    else {
      var excl = wa.exclusions || {};
      misses.sort(function(a,b){ return b.lastPunchDays - a.lastPunchDays; });
      html += '<div class="table-wrap"><table><thead><tr><th>WO#</th><th>Store</th><th>Closed</th><th class="num">Days to close</th><th class="num">Amount</th>' + (admin?'<th>Exclude</th><th>Note</th>':'') + '</tr></thead><tbody>';
      misses.forEach(function(w){
        var ex = excl[w.doc] || { excluded: false, note: '' };
        html += '<tr' + (ex.excluded?' class="bg-bad"':'') + '><td class="mono">' + esc(w.doc) + '</td><td>' + (w.division==='S'?'South':'North') + '</td><td>' + esc(w.closed) + '</td>' +
          '<td class="num">' + (w.lastPunchDays>30?'<span class="cell-bad">':'') + w.lastPunchDays + (w.lastPunchDays>30?'</span>':'') + '</td><td class="num">' + money(w.total) + '</td>' +
          (admin ? '<td style="text-align:center"><input type="checkbox" data-waex="' + esc(w.doc) + '"' + (ex.excluded?' checked':'') + '></td><td><input type="text" data-waexnote="' + esc(w.doc) + '" value="' + esc(ex.note) + '" placeholder="reason…" style="min-width:150px"></td>' : '') + '</tr>';
      });
      html += '</tbody></table></div>';
    }
    html += '</div>';
  }
  sec.innerHTML = html;

  if (admin) {
    var wb = $('#waBonus', sec); if (wb) wb.addEventListener('change', function(){ wa.bonusAmount = num(wb.value); });
    var ws = $('#waSave', sec); if (ws) ws.addEventListener('click', function(){ wa.bonusAmount = num($('#waBonus',sec).value); saveKPI().then(function(){ go('warranty-kpi'); }); });
    sec.querySelectorAll('button[data-wapaid]').forEach(function(btn){
      btn.addEventListener('click', function(){ var k = btn.getAttribute('data-wapaid'); wa.paid = wa.paid || {}; wa.paid[k] = !wa.paid[k]; saveKPI().then(function(){ go('warranty-kpi'); }); });
    });
    function waExRec(d){ wa.exclusions = wa.exclusions || {}; wa.exclusions[d] = wa.exclusions[d] || { excluded:false, note:'' }; return wa.exclusions[d]; }
    sec.querySelectorAll('input[data-waex]').forEach(function(cb){ cb.addEventListener('change', function(){ waExRec(cb.getAttribute('data-waex')).excluded = cb.checked; }); });
    sec.querySelectorAll('input[data-waexnote]').forEach(function(t){ t.addEventListener('change', function(){ waExRec(t.getAttribute('data-waexnote')).note = t.value; }); });
    var wx = $('#waExSave', sec); if (wx) wx.addEventListener('click', function(){ saveKPI().then(function(){ go('warranty-kpi'); }); });
  }
};

/* ============================ 8. ROLE DASHBOARDS ============================ */
RENDER.store = function(sec){
  var st = SESSION.store; if (!st) { sec.innerHTML = '<div class="empty">No store assigned.</div>'; return; }
  var pk = CURRENT_PERIOD;
  var m = metricsFor(pk, st); var s = m ? scoreKPIs(m) : null;
  var html = '<div class="grid g4">';
  html += stat('Store', storeName(st), periodLabel(pk), '');
  html += stat('KPIs Hit', (s?s.points:0) + ' / 4', '', s && s.points>=2 ? 'ok':'bad');
  var pool = techBonusPool(pk, st);
  html += stat('Bonus pool', money(pool.pool), pool.eligible?'eligible':'not eligible', pool.eligible?'ok':'bad');
  var eff = avgEfficiency(pk, st);
  html += stat('Avg Tech Eff', eff==null?'—':pct(eff), 'target '+STATE.kpi.config.techEff+'%', eff!=null&&eff>=STATE.kpi.config.techEff?'ok':'bad');
  html += '</div><div class="spacer"></div>';
  if (s) html += '<div class="card"><h3>KPI Status — ' + periodLabel(pk) + '</h3><div class="spacer"></div>' + kpiChips(s) + '</div><div class="spacer"></div>';

  // store efficiency table (filtered to this division)
  var savedPage = CURRENT_PAGE;
  var effSec = el('div'); RENDER.efficiency(effSec);
  // RENDER.efficiency uses managerStores(); for a manager that is their store already.
  html += effSec.innerHTML;
  // open WO section
  html += renderOpenWOs([st], true);
  sec.innerHTML = html;
  wireOpenWOs(sec);
};

RENDER.tech = function(sec){
  var name = SESSION.name;
  var c = STATE.kpi.config; var target = c.techEff;
  var months = effMonths();
  var mine = STATE.efficiency.byTechMonth[name];
  var html = '<div class="card"><div class="section-title"><h3>' + esc(name) + ' — My Efficiency</h3>' +
    '<span class="muted">target ' + target + '%</span></div>';
  if (!mine) {
    html += '<div class="empty">No efficiency data found for your name yet. It appears once your manager uploads the Labour file.</div>';
  } else {
    var sumB = 0, sumR = 0; var exMap = excludedNBMap();
    html += '<div class="table-wrap"><table><thead><tr><th>Month</th><th class="num">Hours Reported</th><th class="num">Hours Billed</th><th class="num">Efficiency</th></tr></thead><tbody>';
    months.forEach(function(mo){
      var cell = mine[mo];
      var rep = cell ? cell.reported - (exMap[name + '|' + mo] || 0) : 0;
      if (cell && rep > 0) {
        var e = cell.billed / rep * 100; sumB += cell.billed; sumR += rep;
        html += '<tr><td><b>' + esc(mo) + '</b></td><td class="num">' + rep.toFixed(1) + '</td><td class="num">' + cell.billed.toFixed(1) +
          '</td><td class="num ' + (e>=target?'eff-good':'eff-bad') + '">' + e.toFixed(0) + '%</td></tr>';
      } else { html += '<tr><td>' + esc(mo) + '</td><td class="num muted">—</td><td class="num muted">—</td><td class="num muted">—</td></tr>'; }
    });
    var avg = sumR>0 ? sumB/sumR*100 : null;
    html += '</tbody></table></div>';
    html += '<div class="grid g3" style="margin-top:14px">' +
      miniStat('Avg efficiency', avg!=null?pct(avg):'—', avg!=null&&avg>=target) +
      miniStat('Total billed', (sumB).toFixed(1) + ' h', true) +
      miniStat('Total reported', (sumR).toFixed(1) + ' h', true) + '</div>';
  }
  html += '</div><div class="spacer"></div>';

  // flagged WOs for this tech (over 100% and under 75%) — match on the
  // resolved roster name so the tech's own data is found by their login name.
  var mineFlags = consolidateFlagged((STATE.efficiency.flagged || []).filter(function(r){ return (r.display || r.name) === name; })).filter(function(r){ return r.eff < c.effLowFlag; });
  html += '<div class="card"><div class="section-title"><h3>My Flagged Work Orders</h3><span class="muted">under ' + c.effLowFlag + '%</span></div>';
  if (!mineFlags.length) html += '<div class="empty">No flagged work orders. 👍</div>';
  else {
    html += '<div class="table-wrap"><table><thead><tr><th>Month</th><th>WO#</th><th class="num">Punches</th><th class="num">Reported</th><th class="num">Billed</th><th class="num">Eff</th><th>Flag</th></tr></thead><tbody>';
    mineFlags.sort(function(a,b){ return monthIndex(a.month) - monthIndex(b.month); }).forEach(function(r){
      html += '<tr class="bg-bad"><td>' + esc(r.month) + '</td><td class="mono">' + esc(r.docNum) + '</td><td class="num">' + (r.punches||1) + '</td><td class="num">' + r.reported.toFixed(1) +
        '</td><td class="num">' + r.billed.toFixed(1) + '</td><td class="num"><b>' + Math.floor(r.eff) + '%</b></td>' +
        '<td><span class="pill bad">Under ' + c.effLowFlag + '%</span></td></tr>';
    });
    html += '</tbody></table></div>';
  }
  html += '</div><div class="spacer"></div>';

  // this tech's own non-billable hours
  var myNB = (STATE.efficiency.nonBillable || []).filter(function(r){ return r.display === name; });
  var nbTotal = myNB.reduce(function(a,r){ return a + r.hours; }, 0);
  html += '<div class="card"><div class="section-title"><h3>My Non-Billable Hours</h3><span class="muted">' + nbTotal.toFixed(1) + ' hrs total</span></div>';
  if (!myNB.length) html += '<div class="empty">No non-billable hours recorded.</div>';
  else {
    var excl = STATE.kpi.effExclusions || {};
    myNB.sort(function(a,b){ return String(a.month).localeCompare(String(b.month)); });
    html += '<div class="table-wrap"><table><thead><tr><th>Month</th><th>Category</th><th class="num">Entries</th><th class="num">Hours</th><th>Status</th></tr></thead><tbody>';
    myNB.forEach(function(r){
      var exHrs = nbExcludedHours(r);
      html += '<tr' + (exHrs > 0 ? ' class="bg-ok"' : '') + '><td>' + esc(r.month) + '</td><td>' + esc(r.category) + '</td><td class="num">' + (r.count||1) + '</td>' +
        '<td class="num">' + r.hours.toFixed(1) + '</td><td>' + (exHrs > 0 ? '<span class="pill ok">' + exHrs.toFixed(1) + ' hr excl</span>' : '<span class="pill muted">in calc</span>') + '</td></tr>';
    });
    html += '</tbody></table></div>';
  }
  html += '</div>';
  sec.innerHTML = html;
};

RENDER.support = function(sec){
  var queues = SESSION.queues && SESSION.queues.length ? SESSION.queues : (CFG.queuesForStaff(SESSION.name) || []);
  var html = '<div class="card"><div class="section-title"><h3>' + esc(SESSION.name) + ' — My Open WO Queue</h3>' +
    '<span class="muted">' + (queues.length ? queues.join(', ') : 'no queue assigned') + '</span></div>';
  if (!STATE.pos.rows.length) {
    html += '<div class="empty">No open work-order data uploaded yet.</div></div>';
    sec.innerHTML = html; return;
  }
  var rows = STATE.pos.rows.filter(function(r){ return queues.indexOf(r.prefix) !== -1; });
  html += '</div>' + openWOTable(rows, true);
  sec.innerHTML = html;
  wireOpenWOs(sec);
};

/* ---- open WO rendering shared by store dashboard + support ---- */
function renderOpenWOs(stores, withFlag){
  if (!STATE.pos.rows.length) return '<div class="spacer"></div><div class="empty">No open work-order data uploaded yet.</div>';
  var divs = stores.map(storeDivision);
  var rows = STATE.pos.rows.filter(function(r){ return divs.indexOf(r.division) !== -1; });
  return '<div class="spacer"></div><div class="section-title"><h3>Open Work Orders</h3></div>' + openWOTable(rows, withFlag);
}

function openWOTable(rows, withExclude){
  var c = STATE.kpi.config;
  rows = rows.slice().sort(function(a,b){ return b.ageDays - a.ageDays; });
  var totalOpen = rows.reduce(function(a,r){ return STATE.exclusions[r.doc] && STATE.exclusions[r.doc].excluded ? a : a + r.amount; }, 0);
  var html = '<div class="card"><div class="section-title"><h3>Open WO Detail</h3><span class="muted">' + rows.length + ' open · total ' + money(totalOpen) + '</span></div>';
  html += '<div class="table-wrap"><table><thead><tr><th>WO#</th><th>Type</th><th>Sold By</th><th class="num">Amount</th><th class="num">Age</th>' +
    (withExclude ? '<th>Exclude</th><th>Note</th>' : '') + '</tr></thead><tbody>';
  rows.forEach(function(r){
    var rowCls = r.ageDays > c.woCriticalDays ? 'bg-crit' : r.ageDays > c.woWarnDays ? 'bg-warn' : '';
    var ex = STATE.exclusions[r.doc] || { excluded: false, note: '' };
    html += '<tr class="' + rowCls + '"><td class="mono">' + esc(r.doc) + '</td><td>' + esc(r.code || r.prefix) + '</td><td>' + esc(r.soldBy) + '</td>' +
      '<td class="num">' + money(r.amount) + '</td>' +
      '<td class="num">' + r.ageDays + 'd ' + (r.ageDays>c.woCriticalDays?'<span class="pill bad">crit</span>':r.ageDays>c.woWarnDays?'<span class="pill warn">warn</span>':'') + '</td>';
    if (withExclude) {
      html += '<td style="text-align:center"><input type="checkbox" data-ex="' + esc(r.doc) + '"' + (ex.excluded?' checked':'') + '></td>' +
        '<td><input type="text" data-note="' + esc(r.doc) + '" value="' + esc(ex.note) + '" placeholder="note…" style="min-width:160px"></td>';
    }
    html += '</tr>';
  });
  html += '</tbody></table></div>';
  if (withExclude) html += '<div style="text-align:right;margin-top:12px"><button class="btn btn-primary btn-sm" id="exSave">Save exclusions & notes</button></div>';
  html += '</div>';
  return html;
}
function wireOpenWOs(sec){
  sec.querySelectorAll('input[data-ex]').forEach(function(cb){
    cb.addEventListener('change', function(){
      var d = cb.getAttribute('data-ex');
      STATE.exclusions[d] = STATE.exclusions[d] || { excluded:false, note:'' };
      STATE.exclusions[d].excluded = cb.checked;
    });
  });
  sec.querySelectorAll('input[data-note]').forEach(function(t){
    t.addEventListener('change', function(){
      var d = t.getAttribute('data-note');
      STATE.exclusions[d] = STATE.exclusions[d] || { excluded:false, note:'' };
      STATE.exclusions[d].note = t.value;
    });
  });
  var sv = $('#exSave', sec); if (sv) sv.addEventListener('click', saveExclusions);
}

/* ============================ 9. CONFIGURATION + UPLOADS ============================ */
RENDER.config = function(sec){
  var c = STATE.kpi.config;
  var html = '<div class="card"><div class="section-title"><h3>KPI Thresholds</h3>' +
    '<button class="btn btn-primary btn-sm" id="cfgSave">Save thresholds</button></div><div class="form-grid">' +
    cfgField('techEff','Tech efficiency target %', c.techEff) +
    cfgField('comeback','Comeback ceiling %', c.comeback) +
    cfgField('svcGm','Service GM floor %', c.svcGm) +
    cfgField('partsGm','Parts GM floor %', c.partsGm) +
    cfgField('wipMax','Open WIP ceiling %', c.wipMax) +
    cfgField('cap','Quarterly bonus cap $', c.cap) +
    cfgField('techShare','Tech share %', c.techShare) +
    cfgField('growthRate','Growth rate %', c.growthRate) +
    '</div></div><div class="spacer"></div>';

  // three consolidated upload zones
  html += '<div class="card"><div class="section-title"><h3>Consolidated Data Uploads</h3></div>' +
    '<div class="upload-zones">' +
    uploadZone('eff', '⚙', 'Labour Efficiencies', 'TECH hours reported vs billed (.xlsx)', STATE.efficiency.uploadedAt) +
    uploadZone('pos', '📑', 'POS / Open Work Orders', 'Open WO & invoice detail (.xlsx)', STATE.pos.uploadedAt) +
    uploadZone('war', '🛠️', 'Warranty Work Orders', 'Warranty claim export (.xlsx)', STATE.warrantyUpload.uploadedAt) +
    '</div></div>';

  sec.innerHTML = html;

  $('#cfgSave', sec).addEventListener('click', function(){
    ['techEff','comeback','svcGm','partsGm','wipMax','cap','techShare','growthRate'].forEach(function(k){
      c[k] = num($('#cfg_' + k, sec).value);
    });
    saveKPI();
  });

  wireUpload(sec, 'eff', handleEfficiencyFile);
  wireUpload(sec, 'pos', handlePOSFile);
  wireUpload(sec, 'war', handleWarrantyFile);
};
function cfgField(k,label,val){ return '<label class="fld"><span>' + esc(label) + '</span><input type="number" step="any" id="cfg_' + k + '" value="' + val + '"></label>'; }
function uploadZone(id,ic,title,desc,when){
  return '<div class="uz" id="uz_' + id + '"><div class="uz-ic">' + ic + '</div><h4>' + esc(title) + '</h4>' +
    '<p>' + esc(desc) + '</p>' +
    '<button class="btn btn-dark btn-sm" data-pick="' + id + '">Choose file</button>' +
    '<input type="file" accept=".xlsx,.xls,.csv" data-file="' + id + '">' +
    '<div class="uz-when" data-when="' + id + '">' + (when ? 'Last upload: ' + esc(when) : 'No upload yet') + '</div></div>';
}
function wireUpload(sec, id, handler){
  var zone = $('#uz_' + id, sec);
  var input = zone.querySelector('input[type=file]');
  zone.querySelector('button[data-pick]').addEventListener('click', function(){ input.click(); });
  input.addEventListener('change', function(){ if (input.files[0]) processFile(zone, input.files[0], id, handler); });
  ['dragenter','dragover'].forEach(function(ev){ zone.addEventListener(ev, function(e){ e.preventDefault(); zone.classList.add('drag'); }); });
  ['dragleave','drop'].forEach(function(ev){ zone.addEventListener(ev, function(e){ e.preventDefault(); zone.classList.remove('drag'); }); });
  zone.addEventListener('drop', function(e){ var f = e.dataTransfer.files[0]; if (f) processFile(zone, f, id, handler); });
}
function processFile(zone, file, id, handler){
  var whenEl = zone.querySelector('[data-when]');
  whenEl.innerHTML = '<span class="spinner dark"></span> parsing…';
  var reader = new FileReader();
  reader.onload = function(e){
    try {
      var wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      var n = handler(wb);   // handler picks the correct sheet itself
      // also push raw file to Drive for the record
      uploadRaw(file, id);
      whenEl.textContent = 'Loaded ' + n + ' rows · ' + todayStr();
      toast('Parsed ' + n + ' rows from ' + file.name, 'ok');
    } catch (err) {
      whenEl.textContent = 'Parse failed';
      toast('Could not parse: ' + err.message, 'bad');
    }
  };
  reader.readAsArrayBuffer(file);
}

/* Pick the right sheet's rows from a workbook:
 *   1. a sheet whose name matches one of `preferred` (case-insensitive)
 *   2. else the first sheet whose headers contain all `required` columns
 *   3. else the first sheet. */
function pickSheetRows(wb, preferred, required){
  var names = wb.SheetNames || [];
  function rowsOf(nm){ return XLSX.utils.sheet_to_json(wb.Sheets[nm], { defval: '' }); }
  for (var p = 0; p < (preferred||[]).length; p++) {
    for (var i = 0; i < names.length; i++) {
      if (names[i].toLowerCase().trim() === preferred[p].toLowerCase()) return rowsOf(names[i]);
    }
  }
  var norm = function(s){ return String(s).toLowerCase().replace(/[^a-z0-9]/g,''); };
  for (var j = 0; j < names.length; j++) {
    var rows = rowsOf(names[j]);
    if (!rows.length) continue;
    var hdr = Object.keys(rows[0]).map(norm);
    var ok = (required||[]).every(function(req){ var r = norm(req); return hdr.some(function(h){ return h.indexOf(r) !== -1; }); });
    if (ok) return rows;
  }
  return names.length ? rowsOf(names[0]) : [];
}
function uploadRaw(file, id){
  var r = new FileReader();
  r.onload = function(e){
    var b64 = String(e.target.result).split(',')[1] || '';
    var folder = id === 'eff' ? 'Efficiency' : id === 'pos' ? 'POS' : 'Warranty';
    API.uploadFile({ filename: file.name, mimeType: file.type || 'application/octet-stream',
      data: b64, category: folder, period: CURRENT_PERIOD })
      .catch(function(e){ /* non-fatal */ });
  };
  r.readAsDataURL(file);
}

/* ---- column matching helper (tolerant to header naming) ---- */
function pick(row, candidates){
  var keys = Object.keys(row);
  for (var i = 0; i < candidates.length; i++) {
    var want = candidates[i].toLowerCase().replace(/[^a-z0-9]/g,'');
    for (var k = 0; k < keys.length; k++) {
      if (keys[k].toLowerCase().replace(/[^a-z0-9]/g,'') === want) return row[keys[k]];
    }
  }
  // loose contains match
  for (var j = 0; j < candidates.length; j++) {
    var w = candidates[j].toLowerCase().replace(/[^a-z0-9]/g,'');
    for (var kk = 0; kk < keys.length; kk++) {
      if (keys[kk].toLowerCase().replace(/[^a-z0-9]/g,'').indexOf(w) !== -1) return row[keys[kk]];
    }
  }
  return '';
}

function handleEfficiencyFile(wb){
  // Read the flat "DATA" tab (the pivots build off it); fall back to any sheet
  // with the expected columns.
  var rows = pickSheetRows(wb, ['DATA'], ['Payroll Job Code', 'Hours Reported', 'Hours Billed']);
  var out = [];
  rows.forEach(function(r){
    var jobCode = String(pick(r, ['Payroll Job Code','Job Code','JobCode'])).toUpperCase();
    if (jobCode.indexOf('TECH') === -1) return;            // techs only
    var charge  = String(pick(r, ['CHARGE','Charge Type','Charge'])).toUpperCase();
    var reported = num(pick(r, ['Hours Reported','Hrs Reported','Reported']));
    var billed   = num(pick(r, ['Hours Billed','Hrs Billed','Billed']));
    if (reported <= 0 && billed <= 0) return;
    var name = String(pick(r, ['Employee Name','Employee','Name'])).trim();
    var div  = String(pick(r, ['Employee Division','Division','Div'])).trim().toUpperCase().charAt(0);
    var empNum = String(pick(r, ['Employee#','Employee #','Emp#','Employee Number','Emp Number'])).trim();
    var docNum = String(pick(r, ['POS Document#','Document#','Document #','Doc#','Document Number'])).trim();
    var category = String(pick(r, ['STATUS CODE DESC','Status Code Desc','Labour Type','Category'])).trim();
    var comment = String(pick(r, ['Comment','Comments','Note'])).trim();
    // non-billable rows are KEPT (they count toward reported hours / efficiency)
    var nonBillable = charge.replace(/[^A-Z]/g,'').indexOf('NONBILLABLE') !== -1;
    var month = normMonth(pick(r, ['Month','Period']));
    out.push({
      month: month, name: name, empNum: empNum, division: div, docNum: docNum,
      reported: reported, billed: billed, charge: charge, category: category, comment: comment,
      nonBillable: nonBillable,
      eff: reported > 0 ? (billed / reported * 100) : 0,
      key: [empNum || name, month, docNum, category, comment, reported].join('|')
    });
  });
  aggregateEfficiency(out, todayStr());
  // Persist the aggregate inside the backend's `rows` field (it stores a rows
  // array as-is), so efficiency keeps its own storage cell.
  API.saveEfficiency({ rows: [STATE.efficiency], uploadedAt: STATE.efficiency.uploadedAt }).catch(noop);
  return out.length;
}

// Aggregate raw labour rows into the compact, storable structure.
function aggregateEfficiency(rawRows, when){
  var c = STATE.kpi.config;
  var by = {}, techDiv = {}, nbMap = {}, woMap = {}, detail = [];
  (rawRows || []).forEach(function(r){
    if (!r.name && !r.empNum) return;
    var rec = rosterForEmpNum(r.empNum);
    var display = rec ? rec.name : rosterDisplayForLabour(r.name);
    if (r.division && !techDiv[display]) techDiv[display] = r.division;
    var m = by[display] = by[display] || {};
    var cell = m[r.month] = m[r.month] || { reported: 0, billed: 0, nonBillable: 0 };
    cell.billed += r.billed;
    cell.reported += r.reported;                 // full reported incl non-billable
    if (r.nonBillable) {
      cell.nonBillable += r.reported;
      // aggregate non-billable by tech / month / category (keeps storage small)
      var cat = r.category || r.charge || 'NON-BILLABLE';
      var key = [display, r.month, cat].join('|');
      var nb = nbMap[key] || (nbMap[key] = { key: key, display: display, division: r.division,
        month: r.month, category: cat, hours: 0, count: 0 });
      nb.hours += r.reported; nb.count++;
      // individual punch detail (in-memory only — for the drill-down picker)
      detail.push({ catKey: key, display: display, month: r.month, category: cat,
        comment: r.comment, docNum: r.docNum, hours: r.reported });
    } else if (r.docNum) {
      // consolidate billable punches by work order (one row per WO, not per punch)
      var wk = display + '|' + r.month + '|' + r.docNum;
      var wo = woMap[wk] || (woMap[wk] = { display: display, division: r.division, month: r.month,
        docNum: r.docNum, reported: 0, billed: 0, punches: 0 });
      wo.reported += r.reported; wo.billed += r.billed; wo.punches += 1;
    }
  });
  // flag work orders (consolidated) that came in UNDER the efficiency target
  var flagged = [];
  Object.keys(woMap).forEach(function(k){
    var w = woMap[k];
    if (w.reported <= 0) return;
    var eff = w.billed / w.reported * 100;
    if (eff < c.effLowFlag) {
      flagged.push({ display: w.display, division: w.division, month: w.month, docNum: w.docNum,
        reported: w.reported, billed: w.billed, eff: eff, punches: w.punches });
    }
  });
  // lowest efficiency first; keep the worst 120 so the stored blob fits one cell
  flagged.sort(function(a,b){ return a.eff - b.eff; });
  STATE.efficiency = {
    byTechMonth: by, techDiv: techDiv,
    nonBillable: Object.keys(nbMap).map(function(k){ return nbMap[k]; }),
    flagged: flagged.slice(0, 120), uploadedAt: when || STATE.efficiency.uploadedAt
  };
  STATE.nbDetail = detail;   // in-memory only; never persisted
}

// Accept the load response / aggregate / legacy flat rows and normalise.
function ingestEfficiency(data, when){
  var agg = null, rows = null, up = (data && data.uploadedAt) || when;
  if (data) {
    if (data.byTechMonth) agg = data;                                  // bare aggregate
    else if (Array.isArray(data.rows)) {
      if (data.rows[0] && data.rows[0].byTechMonth) agg = data.rows[0]; // new: rows:[aggregate]
      else rows = data.rows;                                           // legacy flat rows
    } else if (Array.isArray(data)) rows = data;
  }
  if (agg) {
    STATE.efficiency = {
      byTechMonth: agg.byTechMonth || {}, techDiv: agg.techDiv || {},
      nonBillable: agg.nonBillable || [], flagged: agg.flagged || [], uploadedAt: up || null
    };
  } else if (rows && rows.length) {
    // legacy: flat rows array (all billable, no non-billable)
    aggregateEfficiency(rows.map(function(r){ return {
      name: r.name, empNum: r.empNum, division: r.division, month: r.month,
      docNum: r.docNum, reported: num(r.reported), billed: num(r.billed),
      category: '', comment: '', nonBillable: false }; }), up);
  }
}
// kept for callers; exclusions are applied at render time so this is a no-op
function rebuildEfficiencyIndex(){}

/* ---- Labour-data name → roster employee linkage --------------------
 * The Labour export and the roster use different name strings for the
 * same person (e.g. "JARED FRANCIA LOBBAN" vs "Jared Lobban", or the
 * nickname "Ed Kindt" vs "RANDOLF E. KINDT"). We resolve each efficiency
 * row to a roster employee so tables show clean roster names and the tech
 * PIN dashboard (which looks up by roster name) finds the right data.
 *
 * Match order, most→least specific:
 *   1. explicit s.effName override (exact, case-insensitive)
 *   2. same first + last token
 *   3. same first INITIAL + last token   (Al↔Alex, Don↔Donald)
 *   4. unique last token across the roster (Ed↔Randolf Kindt)
 * Unmatched names fall back to their raw Labour name (still visible). */
function nameTokens(name){
  return String(name || '').toUpperCase().replace(/[^A-Z\s]/g, ' ').split(/\s+/).filter(Boolean);
}
// Resolve an efficiency row to its roster display name. Primary join is the
// employee number; name-matching is the fallback / bootstrap.
function rosterForEmpNum(empNum){
  empNum = String(empNum || '').trim().toUpperCase();
  if (!empNum) return null;
  for (var i = 0; i < STATE.staff.length; i++) {
    if (String(STATE.staff[i].empNum || '').trim().toUpperCase() === empNum) return STATE.staff[i];
  }
  return null;
}
function rosterDisplayForLabour(labourName){
  var toks = nameTokens(labourName);
  if (!toks.length) return labourName;
  var first = toks[0], last = toks[toks.length - 1];

  var byFirstLast = null, byInitialLast = null, lastMatches = [];
  for (var i = 0; i < STATE.staff.length; i++) {
    var s = STATE.staff[i];
    if (s.effName && s.effName.toUpperCase() === String(labourName).toUpperCase()) return s.name;
    var rt = nameTokens(s.name);
    if (!rt.length) continue;
    var rFirst = rt[0], rLast = rt[rt.length - 1];
    if (rt.length < 2) continue;                 // single-token roster names can't last-match
    if (rLast === last) {
      lastMatches.push(s.name);
      if (rFirst === first) byFirstLast = s.name;
      else if (rFirst.charAt(0) === first.charAt(0)) byInitialLast = s.name;
    }
  }
  if (byFirstLast) return byFirstLast;
  if (byInitialLast) return byInitialLast;
  if (lastMatches.length === 1) return lastMatches[0];
  return labourName;                              // no confident match
}

function handlePOSFile(wb){
  var rows = pickSheetRows(wb, ['DATA','POS','Sheet0'], ['Document Status', 'Document']);
  var out = [];
  rows.forEach(function(r){
    var status = String(pick(r, ['Document Status','Status'])).trim().toUpperCase().charAt(0);
    if (status !== 'O') return;                            // open only
    var doc = String(pick(r, ['Document#','Document #','Doc#','Document Number'])).trim();
    var prefix = doc.slice(0,2).toUpperCase();
    var div = String(pick(r, ['Division','Div'])).trim().toUpperCase().charAt(0);
    if (!div && (prefix === 'WS' || prefix === 'IS')) div = 'S';
    if (!div && (prefix === 'WM' || prefix === 'IM')) div = 'M';
    var opened = parseDate(pick(r, ['Date Opened','Opened','Open Date']));
    out.push({
      division: div, doc: doc, prefix: prefix,
      code: String(pick(r, ['Document Code','Code'])).trim(),
      status: status,
      soldBy: String(pick(r, ['Document Sold By Name','Sold By','Salesperson'])).trim(),
      amount: num(pick(r, ['Document Amount','Document Amount (Total)','Amount','Total'])),
      opened: opened ? opened.toISOString().slice(0,10) : '',
      ageDays: opened ? Math.max(0, Math.round((Date.now() - opened.getTime()) / 86400000)) : 0
    });
  });
  ingestPOS(out, todayStr());
  API.savePOS({ rows: out, uploadedAt: STATE.pos.uploadedAt }).catch(noop);
  return out.length;
}
function ingestPOS(rows, when){
  STATE.pos.rows = rows || [];
  STATE.pos.uploadedAt = when || STATE.pos.uploadedAt;
}

function handleWarrantyFile(wb){
  // Parse the Warranty Work Orders export — the data source for the
  // Warranty Admin KPI (close-within-14-days-of-last-punch).
  var rows = pickSheetRows(wb, ['Sheet0','DATA'], ['Document#', 'WIP Last Punch']);
  var out = [];
  rows.forEach(function(r){
    var doc = String(pick(r, ['Document#','Document #','Doc#','Document Number'])).trim();
    if (!doc) return;
    var type = String(pick(r, ['Document Type (W/S/Q)','Document Type','Type'])).trim().toUpperCase().charAt(0);
    var code = String(pick(r, ['Document Code','Code'])).trim().toUpperCase();
    if (type && type !== 'W' && code !== 'W') return;        // warranty docs only
    var div = String(pick(r, ['Division','Div'])).trim().toUpperCase().charAt(0);
    if (!div) div = doc.slice(1,2).toUpperCase();             // WM.. / WS..
    var closed = parseDate(pick(r, ['Date Closed/Printed','Date Closed','Closed','Date Closed/Printed (CCYYMMDD)']));
    var opened = parseDate(pick(r, ['Date Opened','Opened','Date Opened (CCYYMMDD)']));
    var intFlag = String(pick(r, ['Work Order: Int. flag (A/S)','Int. flag (A/S)','Int Flag'])).trim();
    var warrCode = String(pick(r, ['Work Order: Warr. Addrs#','Warr. Addrs#','Work Order: Warranty Code','Warranty Code'])).toUpperCase();
    out.push({
      doc: doc, division: div, type: type || 'W',
      status: String(pick(r, ['Document Status (O/R/C/k=Cancelled)','Document Status','Status'])).trim().toUpperCase().charAt(0),
      soldBy: String(pick(r, ['Document Sold By Name','Sold By'])).trim(),
      opened: opened ? opened.toISOString().slice(0,10) : '',
      closed: closed ? closed.toISOString().slice(0,10) : '',
      quarter: closed ? quarterOfDate(closed) : '',
      lastPunchDays: num(pick(r, ['WIP Last Punch Days to Close (Closed - Last Punch)','WIP Last Punch Days to Close','Last Punch Days to Close'])),
      daysToClose: num(pick(r, ['Days to Close (Closed-Opened)','Days to Close'])),
      total: num(pick(r, ['Document Amount (Total)','Document Amount','Total'])),
      // internal warranty = the internal warranty account codes (the A/S "Int flag"
      // is an address/stock indicator, not an internal-vs-OEM signal)
      internal: warrCode.indexOf('05101') !== -1 || warrCode.indexOf('S05102') !== -1, intFlag: intFlag
    });
  });
  STATE.warrantyWO = { rows: out, uploadedAt: todayStr() };
  STATE.warrantyUpload.uploadedAt = todayStr();
  API.saveWarranty({ wo: out, uploadedAt: STATE.warrantyWO.uploadedAt }).catch(noop);
  return out.length;
}

// Calendar quarter key 'YYYY-Qn' for a Date.
function quarterOfDate(d){
  if (!d) return '';
  var q = Math.floor(d.getMonth() / 3) + 1;
  return d.getFullYear() + '-Q' + q;
}

/* ============================ 10. MANAGER KPI PDF EXPORT ============================ */
function exportManagerPDF(email){
  var nm = CFG.MANAGER_BONUS[email].name;
  var store = CFG.MANAGER_BONUS[email].store;
  var y = managerYtd(email);
  var rowsHtml = periodsInYear(2026).map(function(p){
    var b = managerBonusFor(email, p.key);
    var tick = function(ok){ return ok ? '<span style="color:#16a34a;font-weight:700">✓</span>' : '<span style="color:#dc2626;font-weight:700">✕</span>'; };
    var m = metricsFor(p.key, store);
    var s = b.score;
    return '<tr>' +
      '<td><b>Q' + p.q + '</b></td>' +
      '<td style="text-align:center">' + tick(s.kpi1) + '<div class="v">' + (m?pct(s.eff):'—') + '</div></td>' +
      '<td style="text-align:center">' + tick(s.kpi2) + '<div class="v">' + (m?pct(m.svcGm):'—') + '</div></td>' +
      '<td style="text-align:center">' + tick(s.kpi3) + '<div class="v">' + (m?pct(m.partsGm):'—') + '</div></td>' +
      '<td style="text-align:center">' + tick(s.kpi4) + '<div class="v">' + (m?pct(s.wipPct):'—') + '</div></td>' +
      '<td style="text-align:center"><b>' + s.points + '/4</b></td>' +
      '<td style="text-align:right"><b>' + money(b.payout) + '</b></td>' +
      '<td style="text-align:center">' + (b.paid ? '<span class="paid">PAID</span>' : '<span class="due">due</span>') + '</td>' +
    '</tr>';
  }).join('');

  var capPct = y.cap ? Math.min(100, y.earned / y.cap * 100) : 0;
  var doc = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>KPI Bonus — ' + esc(nm) + '</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;600;700&family=Barlow+Condensed:wght@600;700&display=swap" rel="stylesheet">' +
    '<style>' +
    'body{font-family:Barlow,sans-serif;color:#1c1c1c;margin:0;padding:0}' +
    '.wrap{max-width:780px;margin:0 auto;padding:28px}' +
    '.hd{background:#1c1c1c;color:#fff;border-radius:12px;padding:22px 26px;display:flex;justify-content:space-between;align-items:center;border-bottom:5px solid #E8620A}' +
    '.hd h1{font-family:Barlow Condensed;margin:0;font-size:26px}.hd p{margin:3px 0 0;color:#bbb;font-size:13px}' +
    '.mark{width:52px;height:52px;border-radius:10px;background:#E8620A;display:flex;align-items:center;justify-content:center;font-family:Barlow Condensed;font-weight:700;font-size:20px}' +
    '.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:18px 0}' +
    '.s{border:1px solid #e6e6e6;border-top:4px solid #E8620A;border-radius:10px;padding:12px 14px}' +
    '.s .l{font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.5px}.s .v{font-family:Barlow Condensed;font-weight:700;font-size:24px}' +
    'table{width:100%;border-collapse:collapse;margin-top:8px;font-size:13px}' +
    'th{background:#1c1c1c;color:#fff;padding:9px;text-align:center;font-family:Barlow Condensed}' +
    'th:first-child,td:first-child{text-align:left}td{padding:9px;border-top:1px solid #eee}.v{font-size:11px;color:#777}' +
    '.paid{background:#eafaf0;color:#16a34a;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700}' +
    '.due{background:#f0f0f0;color:#777;padding:2px 8px;border-radius:99px;font-size:11px}' +
    '.bar{height:12px;background:#eee;border-radius:99px;overflow:hidden;margin-top:6px}.bar i{display:block;height:100%;background:#E8620A}' +
    '.box{border:1px solid #e6e6e6;border-radius:10px;padding:14px 16px;margin-top:18px;background:#fafafa;font-size:13px}' +
    '.box b{font-family:Barlow Condensed}' +
    '@media print{.noprint{display:none}body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}' +
    '</style></head><body><div class="wrap">' +
    '<div class="hd"><div><h1>KPI Incentive Bonus</h1><p>' + esc(nm) + ' · ' + storeName(store) + ' · FY2026</p></div><div class="mark">HPE</div></div>' +
    '<div class="stats">' +
      '<div class="s"><div class="l">Annual cap</div><div class="v">' + money(y.cap) + '</div></div>' +
      '<div class="s"><div class="l">YTD earned</div><div class="v">' + money(y.earned) + '</div></div>' +
      '<div class="s"><div class="l">Paid</div><div class="v">' + money(y.paid) + '</div></div>' +
      '<div class="s"><div class="l">Remaining</div><div class="v">' + money(y.remaining) + '</div></div>' +
    '</div>' +
    '<div style="font-size:12px;color:#777">Progress to annual cap — ' + capPct.toFixed(0) + '%</div>' +
    '<div class="bar"><i style="width:' + capPct + '%"></i></div>' +
    '<table><thead><tr><th>Qtr</th><th>Eff & Comeback</th><th>Service GM</th><th>Parts GM</th><th>Open WIP</th><th>Score</th><th>Payout</th><th>Status</th></tr></thead>' +
    '<tbody>' + rowsHtml + '</tbody></table>' +
    '<div class="box"><b>How this bonus is calculated</b><br>' +
      'Payout each quarter = KPIs hit × (Annual cap ÷ 4 quarters ÷ 4 KPIs). ' +
      'At ' + money(y.cap) + '/yr that is ' + money(y.cap/4) + ' maximum per quarter and ' + money(y.cap/16) + ' per KPI achieved. ' +
      'Year-end growth bonus is tracked separately.</div>' +
    '<div class="noprint" style="text-align:right;margin-top:20px"><button onclick="window.print()" style="background:#E8620A;color:#fff;border:0;border-radius:8px;padding:11px 20px;font-weight:700;font-family:Barlow;cursor:pointer">Print / Save PDF</button></div>' +
    '</div></body></html>';

  var w = window.open('', '_blank');
  if (!w) { toast('Allow pop-ups to export the PDF.', 'bad'); return; }
  w.document.open(); w.document.write(doc); w.document.close();
}

/* ============================ SHARED UTILITIES ============================ */
function managerStores(){
  if (SESSION.role === 'manager' && SESSION.store) return [SESSION.store];
  return ['south', 'north'];
}
function managerFilterStaff(list){
  if (SESSION.role === 'manager' && SESSION.store) return list.filter(function(s){ return s.store === SESSION.store; });
  return list;
}
function hasEfficiency(){ return Object.keys(STATE.efficiency.byTechMonth || {}).length > 0; }
function effTechsForDivision(div){
  var by = STATE.efficiency.byTechMonth || {}, td = STATE.efficiency.techDiv || {};
  return Object.keys(by).filter(function(name){ return td[name] === div; }).sort();
}
function effMonths(){
  var set = {}, by = STATE.efficiency.byTechMonth || {};
  Object.keys(by).forEach(function(n){ Object.keys(by[n]).forEach(function(m){ if (m) set[m] = true; }); });
  var order = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return Object.keys(set).sort(function(a,b){ return order.indexOf(a) - order.indexOf(b); });
}
// Hours removed from a non-billable category (entered directly or via the punch
// picker); legacy {excluded:true} means "exclude all hours of that category".
function nbExcludedHours(nb){
  var ex = (STATE.kpi.effExclusions || {})[nb.key];
  if (!ex) return 0;
  var hrs = (ex.excludeHours != null) ? num(ex.excludeHours) : (ex.excluded ? nb.hours : 0);
  return Math.max(0, Math.min(hrs, nb.hours));
}
// Excluded non-billable hours keyed by "display|month".
function excludedNBMap(){
  var map = {};
  (STATE.efficiency.nonBillable || []).forEach(function(nb){
    var hrs = nbExcludedHours(nb);
    if (hrs > 0) { var k = nb.display + '|' + nb.month; map[k] = (map[k] || 0) + hrs; }
  });
  return map;
}
function avgEfficiency(periodKey, store){
  // store === null → all stores. Efficiency denominator excludes manually-
  // excluded non-billable hours.
  var by = STATE.efficiency.byTechMonth || {}, td = STATE.efficiency.techDiv || {}, exMap = excludedNBMap();
  var b = 0, r = 0;
  Object.keys(by).forEach(function(name){
    if (store && divisionStore(td[name]) !== store) return;
    Object.keys(by[name]).forEach(function(m){
      var cell = by[name][m]; b += cell.billed;
      r += cell.reported - (exMap[name + '|' + m] || 0);
    });
  });
  return r > 0 ? (b / r * 100) : null;
}
var MONTHS_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function monthIndex(m){ var i = MONTHS_ABBR.indexOf(m); return i === -1 ? 99 : i; }
function normMonth(v){
  if (v == null || v === '') return '';
  if (v instanceof Date) return MONTHS_ABBR[v.getMonth()];
  var s = String(v).trim();
  // YYYYMM (e.g. 202601) -> month from the last two digits
  if (/^\d{6}$/.test(s)) { var mm = parseInt(s.slice(4, 6), 10); if (mm >= 1 && mm <= 12) return MONTHS_ABBR[mm - 1]; }
  // YYYY-MM or MM/YYYY etc.
  var m6 = s.match(/(\d{4})[-/](\d{1,2})/); if (m6) { var mo = parseInt(m6[2],10); if (mo>=1&&mo<=12) return MONTHS_ABBR[mo-1]; }
  // leading 1-2 digit month
  var m = s.match(/^(\d{1,2})\b/);
  if (m && Number(m[1]) >= 1 && Number(m[1]) <= 12 && !/[a-z]/i.test(s)) return MONTHS_ABBR[Number(m[1]) - 1];
  return s.slice(0, 3).replace(/^./, function(c){ return c.toUpperCase(); });
}
function parseDate(v){
  if (!v) return null;
  if (v instanceof Date) return v;
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
function todayStr(){
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

/* ---- tiny modal ---- */
function modal(title, bodyHtml, onMount){
  closeModal();
  var ov = el('div', { id: '__modal', style: 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:60;display:flex;align-items:center;justify-content:center;padding:20px' });
  var box = el('div', { class: 'card', style: 'max-width:640px;width:100%;max-height:86vh;overflow:auto' });
  box.innerHTML = '<div class="section-title"><h3>' + esc(title) + '</h3><button class="btn btn-ghost btn-sm" id="__mclose">Close</button></div>' + bodyHtml;
  ov.appendChild(box); document.body.appendChild(ov);
  $('#__mclose', box).addEventListener('click', closeModal);
  ov.addEventListener('click', function(e){ if (e.target === ov) closeModal(); });
  if (onMount) onMount(box);
}
function closeModal(){ var m = $('#__modal'); if (m) m.parentNode.removeChild(m); }

/* ============================ BOOTSTRAP ============================ */
function buildPeriodPicker(){
  var sel = $('#qSel'); sel.innerHTML = '';
  PERIODS.forEach(function(p){ sel.appendChild(el('option', { value: p.key }, p.label)); });
  // provisional default (refined to latest-with-data after loadAll)
  CURRENT_PERIOD = (PERIODS[1] || PERIODS[0]).key;
  sel.value = CURRENT_PERIOD;
  sel.addEventListener('change', function(){ CURRENT_PERIOD = sel.value; if (CURRENT_PAGE) go(CURRENT_PAGE); });
}

// Latest period that actually has entered quarter data (for either store).
function latestPeriodWithData(){
  for (var i = PERIODS.length - 1; i >= 0; i--) {
    var q = STATE.kpi.quarters[PERIODS[i].key];
    if (q && (q.south || q.north)) return PERIODS[i].key;
  }
  return null;
}
function applyDefaultPeriod(){
  var def = latestPeriodWithData();
  if (def) {
    CURRENT_PERIOD = def;
    var sel = $('#qSel'); if (sel) sel.value = def;
  }
}

function init(){
  $('#sbUser').textContent = SESSION.name || SESSION.email || '—';
  $('#sbRole').textContent = (SESSION.role || '').toUpperCase() + (SESSION.store ? ' · ' + storeName(SESSION.store) : '');
  $('#signout').addEventListener('click', function(e){ e.preventDefault(); AUTH.logout('../index.html'); });
  $('#reloadBtn').addEventListener('click', function(){ loadAll().then(function(){ if (CURRENT_PAGE) go(CURRENT_PAGE); }); });
  buildPeriodPicker();
  buildNav();
  loadAll().then(function(){ applyDefaultPeriod(); go(landingPage()); });
}

init();

})();
