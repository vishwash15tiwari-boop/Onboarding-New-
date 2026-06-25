// ═══════════════════════════════════════════════════════
//  ONBOARDING CONTROL TOWER — Apps Script Backend
//  Sheet ID: 1DNXJnjUsrcco9eq073tWIZOhYYyXmp4ajKxyGkpj8NU
// ═══════════════════════════════════════════════════════

var SHEET_ID = '1DNXJnjUsrcco9eq073tWIZOhYYyXmp4ajKxyGkpj8NU';
var SLA_DAYS  = 30; // default SLA threshold

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('dashboard');
  template.payload = JSON.stringify(buildDashboardPayload());
  return template.evaluate()
    .setTitle('Onboarding Control Tower')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ── Read & normalise sheet rows ───────────────────────
function getRawCases() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var hdrs = data[0].map(function(h){ return String(h).trim().toLowerCase().replace(/\s+/g,'_'); });

  return data.slice(1).map(function(row, idx){
    var c = {};
    hdrs.forEach(function(h, i){ c[h] = row[i]; });

    // Normalise dates
    ['created_date','onboarded_date','firstshipmentdate','status_changed_on',
     'review_submission_date','level1','level2','level3','level4',
     'level1_rejected1','level1_rejected2','level2_rejected1','level2_rejected2',
     'level3_rejected1','level3_rejected2','level4_rejected1','level4_rejected2'
    ].forEach(function(f){
      if (c[f] instanceof Date && !isNaN(c[f])) {
        c[f] = Utilities.formatDate(c[f], Session.getScriptTimeZone(), 'dd MMM yyyy');
      } else if (c[f] && String(c[f]).trim() === '') {
        c[f] = '';
      }
    });

    c._rowIdx  = idx;
    c._tatDays = tatDays(c.created_date, c.onboarded_date);
    c._hasL1   = !!c.level1;
    c._hasL2   = !!c.level2;
    c._hasL3   = !!c.level3;
    c._hasL4   = !!c.level4;
    c._stage   = deriveStage(c);
    c._ltv     = parseFloat(String(c.lifetimevalue).replace(/[^0-9.]/g,'')) || 0;

    // Days since created (for open/in-progress cases)
    c._agedays = tatDays(c.created_date, new Date().toISOString());

    // SLA status (only meaningful for active/non-terminal cases)
    c._slaStatus = deriveSLA(c);

    // Days in current stage (status_changed_on → today, or created → today)
    var stageFrom = c.status_changed_on || c.created_date;
    c._stageDays = tatDays(stageFrom, new Date().toISOString());

    return c;
  }).filter(function(c){ return c.id || c.business_name; });
}

// ── Derive current pipeline stage ───────────────────
// Correct flow: Lead → In Progress → In Review (committee) → L1 Approval → L2 Approval → L3 Approval → Onboarded
function deriveStage(c) {
  var cs = String(c.current_status || '').toUpperCase();
  if (cs === 'ONBOARDED')   return 'Onboarded';
  if (cs === 'REACTIVATED') return 'Reactivated';
  if (cs === 'CHURNED')     return 'Churned';
  if (cs === 'DEACTIVATED') return 'Deactivated';
  if (cs === 'SENT_BACK' || cs === 'SENTBACK') return 'Sent Back';
  if (cs === 'REJECTED')    return 'Rejected';
  if (c._hasL4) return 'Onboarded';
  if (c._hasL3) return 'L3 Approval';
  if (c._hasL2) return 'L2 Approval';   // after L2 approved, awaiting L3
  if (c._hasL1) return 'L1 Approval';   // after committee, L1 approved, awaiting L2
  if (c.review_submission_date) return 'In Review';  // docs submitted; committee reviewing
  if (cs === 'IN_PROGRESS' || cs === 'INPROGRESS') return 'In Progress';
  return 'Lead';
}

// ── SLA status ────────────────────────────────────────
function deriveSLA(c) {
  var terminal = ['ONBOARDED','REACTIVATED','CHURNED','DEACTIVATED','REJECTED'];
  var cs = String(c.current_status || '').toUpperCase();
  if (terminal.indexOf(cs) !== -1) return 'Completed';
  var age = c._agedays;
  if (age === null) return 'Unknown';
  if (age > SLA_DAYS)      return 'Overdue';
  if (age === SLA_DAYS)    return 'Due Today';
  if (age >= SLA_DAYS - 5) return 'At Risk';
  return 'On Track';
}

// ── TAT helpers ──────────────────────────────────────
function tatDays(from, to) {
  try {
    var d1 = new Date(from), d2 = new Date(to);
    if (isNaN(d1) || isNaN(d2)) return null;
    return Math.round((d2 - d1) / 86400000);
  } catch(e) { return null; }
}

function avg(arr) {
  var nums = arr.filter(function(x){ return x !== null && !isNaN(x); });
  if (!nums.length) return 0;
  return Math.round(nums.reduce(function(a,b){return a+b;},0) / nums.length);
}

// ── Build full dashboard payload ─────────────────────
function buildDashboardPayload() {
  var cases = getRawCases();
  var total = cases.length;

  // ── Status map ────────────────────────────────────
  var byStatus = {};
  cases.forEach(function(c){
    var s = String(c.current_status || 'UNKNOWN').toUpperCase();
    byStatus[s] = (byStatus[s] || 0) + 1;
  });

  var tats       = cases.map(function(c){ return c._tatDays; });
  var validTATs  = tats.filter(function(t){ return t !== null && t >= 0; });

  // ── Pipeline funnel counts ────────────────────────
  var lvlCounts = {
    registered  : total,
    inReview    : cases.filter(function(c){ return !!c.review_submission_date; }).length,
    l1Done      : cases.filter(function(c){ return c._hasL1; }).length,
    l2Done      : cases.filter(function(c){ return c._hasL2; }).length,
    l3Done      : cases.filter(function(c){ return c._hasL3; }).length,
    l4Done      : cases.filter(function(c){ return c._hasL4; }).length,
    onboarded   : cases.filter(function(c){ return !!c.onboarded_date; }).length
  };

  // ── KPI extras ────────────────────────────────────
  var activeCases = cases.filter(function(c){
    var cs = String(c.current_status||'').toUpperCase();
    return ['ONBOARDED','CHURNED','DEACTIVATED','REACTIVATED'].indexOf(cs) === -1;
  });
  var pendingL1 = activeCases.filter(function(c){ return !c._hasL1 && c.review_submission_date; }).length;
  var pendingL2 = activeCases.filter(function(c){ return c._hasL1 && !c._hasL2; }).length;
  var pendingL3 = activeCases.filter(function(c){ return c._hasL2 && !c._hasL3; }).length;

  // SLA counts (only active)
  var slaMap = { onTrack:0, dueToday:0, atRisk:0, overdue:0 };
  activeCases.forEach(function(c){
    if (c._slaStatus === 'On Track')  slaMap.onTrack++;
    if (c._slaStatus === 'Due Today') slaMap.dueToday++;
    if (c._slaStatus === 'At Risk')   slaMap.atRisk++;
    if (c._slaStatus === 'Overdue')   slaMap.overdue++;
  });

  // ── Category breakdown ────────────────────────────
  var catMap = {};
  cases.forEach(function(c){
    var cat = String(c.business_category || 'Other').trim() || 'Other';
    if (!catMap[cat]) catMap[cat] = { total:0, onboarded:0, churned:0, deactivated:0, tats:[], ltv:0, active:0 };
    catMap[cat].total++;
    var cs = String(c.current_status||'').toUpperCase();
    if (cs==='ONBOARDED'||cs==='REACTIVATED') catMap[cat].onboarded++;
    if (cs==='CHURNED')     catMap[cat].churned++;
    if (cs==='DEACTIVATED') catMap[cat].deactivated++;
    if (['ONBOARDED','CHURNED','DEACTIVATED','REACTIVATED'].indexOf(cs) === -1) catMap[cat].active++;
    if (c._tatDays !== null) catMap[cat].tats.push(c._tatDays);
    catMap[cat].ltv += c._ltv;
  });
  var categories = Object.keys(catMap).sort(function(a,b){
    return catMap[b].total - catMap[a].total;
  }).map(function(k){
    var d = catMap[k];
    return { name:k, total:d.total, onboarded:d.onboarded, churned:d.churned,
             deactivated:d.deactivated, active:d.active, avgTAT:avg(d.tats), ltv:Math.round(d.ltv) };
  });

  // ── Vertical breakdown (Business Vertical: Marketplace, Open Marketplace, EPR…) ──
  var vertMap = {};
  cases.forEach(function(c){
    var vt = String(c.business_vertical || c.vertical || c.vendor_vertical || 'Other').trim() || 'Other';
    if (!vertMap[vt]) vertMap[vt] = { total:0, onboarded:0, churned:0, deactivated:0, active:0, tats:[], ltv:0 };
    vertMap[vt].total++;
    var cs = String(c.current_status||'').toUpperCase();
    if (cs==='ONBOARDED'||cs==='REACTIVATED') vertMap[vt].onboarded++;
    if (cs==='CHURNED')     vertMap[vt].churned++;
    if (cs==='DEACTIVATED') vertMap[vt].deactivated++;
    if (['ONBOARDED','CHURNED','DEACTIVATED','REACTIVATED'].indexOf(cs) === -1) vertMap[vt].active++;
    if (c._tatDays !== null) vertMap[vt].tats.push(c._tatDays);
    vertMap[vt].ltv += c._ltv;
  });
  var verticals = Object.keys(vertMap).sort(function(a,b){
    return vertMap[b].total - vertMap[a].total;
  }).map(function(k){
    var d = vertMap[k];
    return { name:k, total:d.total, onboarded:d.onboarded, churned:d.churned,
             deactivated:d.deactivated, active:d.active, avgTAT:avg(d.tats), ltv:Math.round(d.ltv) };
  });

  // ── State breakdown ───────────────────────────────
  var stateMap = {};
  cases.forEach(function(c){
    var st = String(c.state||'').trim() || 'Unknown';
    stateMap[st] = (stateMap[st]||0) + 1;
  });
  var states = Object.keys(stateMap)
    .filter(function(s){ return s !== 'Unknown'; })
    .sort(function(a,b){ return stateMap[b]-stateMap[a]; })
    .slice(0,10)
    .map(function(s){ return {name:s, count:stateMap[s]}; });

  // ── TAT distribution ──────────────────────────────
  var tatBuckets = { '0-7':0, '8-15':0, '16-30':0, '31-60':0, '60+':0 };
  validTATs.forEach(function(t){
    if (t<=7)       tatBuckets['0-7']++;
    else if (t<=15) tatBuckets['8-15']++;
    else if (t<=30) tatBuckets['16-30']++;
    else if (t<=60) tatBuckets['31-60']++;
    else            tatBuckets['60+']++;
  });

  // ── Monthly trend ─────────────────────────────────
  var monthlyMap = {};
  cases.forEach(function(c){
    if (!c.created_date) return;
    try {
      var d = new Date(c.created_date);
      if (isNaN(d)) return;
      var key = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
      if (!monthlyMap[key]) monthlyMap[key] = { total:0, onboarded:0 };
      monthlyMap[key].total++;
      if (c.onboarded_date) monthlyMap[key].onboarded++;
    } catch(e){}
  });
  var monthlyTrend = Object.keys(monthlyMap).sort().slice(-12).map(function(k){
    return { month:k, total:monthlyMap[k].total, onboarded:monthlyMap[k].onboarded };
  });

  // ── Top cases by LTV ──────────────────────────────
  var topLTV = cases
    .filter(function(c){ return c._ltv > 0; })
    .sort(function(a,b){ return b._ltv - a._ltv; })
    .slice(0,10)
    .map(function(c){
      return { id:c.id, name:c.business_name, category:c.business_category, ltv:c._ltv, status:c.current_status };
    });

  // ── Pending action queue ──────────────────────────
  var pendingActions = {
    awaitingDocs      : activeCases.filter(function(c){ return !c.review_submission_date; }).length,
    awaitingL1        : pendingL1,
    awaitingL2        : pendingL2,
    awaitingL3        : pendingL3,
    sentBack          : cases.filter(function(c){
                          var cs = String(c.current_status||'').toUpperCase();
                          return cs === 'SENT_BACK' || cs === 'SENTBACK';
                        }).length,
    tatBreached       : slaMap.overdue
  };

  // ── Stage-wise breakdown ──────────────────────────
  var stageMap = {};
  cases.forEach(function(c){
    var st = c._stage || 'Unknown';
    if (!stageMap[st]) stageMap[st] = { count:0, stageDays:[] };
    stageMap[st].count++;
    if (c._stageDays !== null && c._stageDays >= 0) stageMap[st].stageDays.push(c._stageDays);
  });
  var stageBreakdown = Object.keys(stageMap).map(function(k){
    return { stage:k, count:stageMap[k].count, avgDays:avg(stageMap[k].stageDays) };
  });

  // ── Cases for table (limit 1000) ──────────────────
  var tableCases = cases.slice(0, 1000).map(function(c){
    return {
      id               : c.id,
      name             : String(c.business_name||c.party_name||'').trim(),
      gstin            : String(c.gstin_number||c.gstin||'').trim(),
      city             : String(c.city||'').trim(),
      state            : String(c.state||'').trim(),
      category         : String(c.business_category||c.category||'').trim(),
      businessVertical : String(c.business_vertical||c.vertical||c.vendor_vertical||'').trim(),
      vendorType       : String(c.vendor_type||c.customer_vendor_type||'').trim(),
      partyType        : String(c.party_type||'').trim(),
      status           : String(c.status||'').trim(),
      currentStatus    : String(c.current_status||'').trim(),
      stage          : c._stage,
      createdDate    : String(c.created_date||'').trim(),
      onboardedDate  : String(c.onboarded_date||'').trim(),
      reviewDate     : String(c.review_submission_date||'').trim(),
      statusChangedOn: String(c.status_changed_on||'').trim(),
      tatDays        : c._tatDays,
      ageDays        : c._agedays,
      stageDays      : c._stageDays,
      slaStatus      : c._slaStatus,
      level1         : String(c.level1||'').trim(),
      level2         : String(c.level2||'').trim(),
      level3         : String(c.level3||'').trim(),
      level4         : String(c.level4||'').trim(),
      ltv            : c._ltv,
      firstShipment  : String(c.firstshipmentdate||'').trim(),
      l1r1           : String(c.level1_rejected1||'').trim(),
      l1r2           : String(c.level1_rejected2||'').trim(),
      l2r1           : String(c.level2_rejected1||'').trim(),
      l2r2           : String(c.level2_rejected2||'').trim(),
      l3r1           : String(c.level3_rejected1||'').trim(),
      l3r2           : String(c.level3_rejected2||'').trim(),
      l4r1           : String(c.level4_rejected1||'').trim(),
      l4r2           : String(c.level4_rejected2||'').trim(),
      contactPerson  : String(c.contact_person||c.contact_name||'').trim(),
      remarks        : String(c.remarks||c.notes||'').trim(),
    };
  });

  return {
    generatedAt     : new Date().toLocaleString('en-IN'),
    total           : total,
    byStatus        : byStatus,
    avgTAT          : avg(validTATs),
    sla             : slaMap,
    pipeline        : lvlCounts,
    pendingActions  : pendingActions,
    stageBreakdown  : stageBreakdown,
    categories      : categories,
    verticals       : verticals,
    states          : states,
    tatBuckets      : tatBuckets,
    monthlyTrend    : monthlyTrend,
    topLTV          : topLTV,
    cases           : tableCases
  };
}
