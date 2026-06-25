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

// ── Flexible field reader ─────────────────────────────
// Tries multiple possible column-name variants, returns first non-empty value.
function pick(c /*, ...keys */) {
  for (var i = 1; i < arguments.length; i++) {
    var v = c[arguments[i]];
    if (v !== undefined && v !== null && String(v).trim() !== '') return v;
  }
  return '';
}

// ── Read & normalise sheet rows ───────────────────────
function getRawCases() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { cases: [], headers: [] };

  // Normalise headers: trim, lowercase, replace spaces/hyphens/slashes with underscore,
  // remove brackets and other noise so "Party ID", "party-id", "party id" all become "party_id"
  var hdrs = data[0].map(function(h){
    return String(h).trim().toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')   // any non-alphanum run → _
      .replace(/^_+|_+$/g, '');       // strip leading/trailing _
  });

  var rows = data.slice(1).map(function(row, idx){
    var c = {};
    hdrs.forEach(function(h, i){ c[h] = row[i]; });

    // Normalise ALL Date cells automatically — no need to hardcode field names
    hdrs.forEach(function(h){
      if (c[h] instanceof Date && !isNaN(c[h])) {
        c[h] = Utilities.formatDate(c[h], Session.getScriptTimeZone(), 'dd MMM yyyy');
      }
    });

    // Also handle the specific date keys we know about (belt-and-suspenders)
    [
      'created_date','creation_date','registration_date','onboarded_date','activation_date',
      'firstshipmentdate','first_shipment_date','first_shipment',
      'status_changed_on','status_change_date',
      'review_submission_date','docs_submission_date','submission_date',
      'level1','level1_date','l1_date','l1_approval_date',
      'level2','level2_date','l2_date','l2_approval_date',
      'level3','level3_date','l3_date','l3_approval_date',
      'level4','level4_date','l4_date','l4_approval_date',
      'level1_rejected1','level1_rejected2','level1_rejection1','level1_rejection2',
      'level2_rejected1','level2_rejected2','level2_rejection1','level2_rejection2',
      'level3_rejected1','level3_rejected2','level3_rejection1','level3_rejection2',
      'level4_rejected1','level4_rejected2','level4_rejection1','level4_rejection2'
    ].forEach(function(f){
      if (c[f] instanceof Date && !isNaN(c[f])) {
        c[f] = Utilities.formatDate(c[f], Session.getScriptTimeZone(), 'dd MMM yyyy');
      }
    });

    c._rowIdx = idx;

    // ── Resolve canonical fields using aliases ────────
    // ID
    var id = pick(c,
      'id','party_id','case_id','vendor_id','customer_id','onboarding_id',
      'party_no','case_no','vendor_no','sl_no','sno','sr_no','s_no');
    c._id = String(id).trim();

    // Name
    var name = pick(c,
      'business_name','party_name','vendor_name','name','company_name',
      'firm_name','organisation_name','organization_name','entity_name');
    c._name = String(name).trim();

    // GSTIN
    var gstin = pick(c,
      'gstin','gstin_number','gst_number','gst_no','gstin_no','gst');
    c._gstin = String(gstin).trim();

    // Business category
    var cat = pick(c,
      'business_category','category','waste_category','product_category',
      'service_category','industry','sector');
    c._category = String(cat).trim();

    // Business vertical
    var vert = pick(c,
      'business_vertical','vertical','vendor_vertical','biz_vertical',
      'business_type','business_segment','segment','division');
    c._businessVertical = String(vert).trim();

    // Vendor / customer type
    var vtype = pick(c,
      'vendor_type','customer_vendor_type','customer_type','party_type_2',
      'vendor_category','entity_type','stakeholder_type');
    c._vendorType = String(vtype).trim();

    // Party type
    var ptype = pick(c,
      'party_type','partner_type','account_type','type');
    c._partyType = String(ptype).trim();

    // Current status
    var cstatus = pick(c,
      'current_status','status','onboarding_status','vendor_status',
      'account_status','application_status','stage_status');
    c._currentStatus = String(cstatus).trim();

    // Dates
    c._createdDate = String(pick(c,
      'created_date','creation_date','registration_date','registered_date',
      'created_on','date_of_registration','doc','date_of_creation') || '').trim();

    c._onboardedDate = String(pick(c,
      'onboarded_date','activation_date','onboarding_date','activated_date',
      'go_live_date','live_date','onboarded_on') || '').trim();

    c._reviewDate = String(pick(c,
      'review_submission_date','docs_submission_date','submission_date',
      'document_submission_date','review_date','submitted_date',
      'docs_submitted_date','documents_submitted_date') || '').trim();

    c._statusChangedOn = String(pick(c,
      'status_changed_on','status_change_date','status_updated_on',
      'last_status_change','status_change_on') || '').trim();

    c._firstShipment = String(pick(c,
      'firstshipmentdate','first_shipment_date','first_shipment',
      'first_txn_date','first_transaction_date') || '').trim();

    // L1-L4 approval dates
    c._level1 = String(pick(c,'level1','level1_date','l1_date','l1_approval_date','l1') || '').trim();
    c._level2 = String(pick(c,'level2','level2_date','l2_date','l2_approval_date','l2') || '').trim();
    c._level3 = String(pick(c,'level3','level3_date','l3_date','l3_approval_date','l3') || '').trim();
    c._level4 = String(pick(c,'level4','level4_date','l4_date','l4_approval_date','l4') || '').trim();

    // Rejection dates
    c._l1r1 = String(pick(c,'level1_rejected1','level1_rejection1','l1_rej1','l1_rejected_1') || '').trim();
    c._l1r2 = String(pick(c,'level1_rejected2','level1_rejection2','l1_rej2','l1_rejected_2') || '').trim();
    c._l2r1 = String(pick(c,'level2_rejected1','level2_rejection1','l2_rej1','l2_rejected_1') || '').trim();
    c._l2r2 = String(pick(c,'level2_rejected2','level2_rejection2','l2_rej2','l2_rejected_2') || '').trim();
    c._l3r1 = String(pick(c,'level3_rejected1','level3_rejection1','l3_rej1','l3_rejected_1') || '').trim();
    c._l3r2 = String(pick(c,'level3_rejected2','level3_rejection2','l3_rej2','l3_rejected_2') || '').trim();
    c._l4r1 = String(pick(c,'level4_rejected1','level4_rejection1','l4_rej1','l4_rejected_1') || '').trim();
    c._l4r2 = String(pick(c,'level4_rejected2','level4_rejection2','l4_rej2','l4_rejected_2') || '').trim();

    // Location
    c._city  = String(pick(c,'city','vendor_city','town','district') || '').trim();
    c._state = String(pick(c,'state','vendor_state','state_name','province') || '').trim();

    // Contact
    c._contact = String(pick(c,
      'contact_person','contact_name','poc_name','point_of_contact',
      'primary_contact','contact') || '').trim();

    // Remarks
    c._remarks = String(pick(c,'remarks','notes','comments','remark','note') || '').trim();

    // LTV
    var ltvRaw = pick(c,
      'lifetimevalue','lifetime_value','ltv','life_time_value',
      'total_revenue','revenue','gmv','total_gmv');
    c._ltv = parseFloat(String(ltvRaw).replace(/[^0-9.]/g,'')) || 0;

    // ── Derived booleans ──────────────────────────────
    c._hasL1 = !!c._level1;
    c._hasL2 = !!c._level2;
    c._hasL3 = !!c._level3;
    c._hasL4 = !!c._level4;

    // ── TAT & SLA ─────────────────────────────────────
    c._tatDays   = tatDays(c._createdDate, c._onboardedDate);
    c._agedays   = tatDays(c._createdDate, new Date().toISOString());
    c._slaStatus = deriveSLA(c);

    var stageFrom = c._statusChangedOn || c._createdDate;
    c._stageDays = tatDays(stageFrom, new Date().toISOString());

    c._stage = deriveStage(c);

    return c;
  });

  // Filter: keep rows that have at least an ID or a name
  var cases = rows.filter(function(c){
    return c._id || c._name;
  });

  return { cases: cases, headers: hdrs };
}

// ── Derive current pipeline stage ───────────────────
// Correct flow: Lead → In Progress → In Review (committee) → L1 Approval → L2 Approval → L3 Approval → Onboarded
function deriveStage(c) {
  var cs = String(c._currentStatus || '').toUpperCase().replace(/\s+/g,'_');
  if (cs === 'ONBOARDED')   return 'Onboarded';
  if (cs === 'REACTIVATED') return 'Reactivated';
  if (cs === 'CHURNED')     return 'Churned';
  if (cs === 'DEACTIVATED') return 'Deactivated';
  if (cs === 'SENT_BACK' || cs === 'SENTBACK') return 'Sent Back';
  if (cs === 'REJECTED')    return 'Rejected';
  if (c._hasL4) return 'Onboarded';
  if (c._hasL3) return 'L3 Approval';
  if (c._hasL2) return 'L2 Approval';
  if (c._hasL1) return 'L1 Approval';
  if (c._reviewDate) return 'In Review';
  if (cs === 'IN_PROGRESS' || cs === 'INPROGRESS') return 'In Progress';
  return 'Lead';
}

// ── SLA status ────────────────────────────────────────
function deriveSLA(c) {
  var terminal = ['ONBOARDED','REACTIVATED','CHURNED','DEACTIVATED','REJECTED'];
  var cs = String(c._currentStatus || '').toUpperCase().replace(/\s+/g,'_');
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
    if (!from || !to) return null;
    // Parse "dd MMM yyyy" format explicitly
    var d1 = parseDate(from), d2 = parseDate(to);
    if (!d1 || !d2 || isNaN(d1) || isNaN(d2)) return null;
    return Math.round((d2 - d1) / 86400000);
  } catch(e) { return null; }
}

var MONTHS_MAP = {
  jan:0,feb:1,mar:2,apr:3,may:4,jun:5,
  jul:6,aug:7,sep:8,oct:9,nov:10,dec:11
};
function parseDate(str) {
  if (!str) return null;
  var s = String(str).trim();
  // Try "dd MMM yyyy" or "d MMM yyyy"
  var m = s.match(/^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})$/);
  if (m) {
    var mo = MONTHS_MAP[m[2].toLowerCase()];
    if (mo !== undefined) return new Date(parseInt(m[3]), mo, parseInt(m[1]));
  }
  // Try ISO or other standard formats
  var d = new Date(s);
  return isNaN(d) ? null : d;
}

function avg(arr) {
  var nums = arr.filter(function(x){ return x !== null && !isNaN(x); });
  if (!nums.length) return 0;
  return Math.round(nums.reduce(function(a,b){return a+b;},0) / nums.length);
}

// ── Build full dashboard payload ─────────────────────
function buildDashboardPayload() {
  var result = getRawCases();
  var cases  = result.cases;
  var hdrs   = result.headers;
  var total  = cases.length;

  // ── Status map ────────────────────────────────────
  var byStatus = {};
  cases.forEach(function(c){
    var s = String(c._currentStatus || 'UNKNOWN').toUpperCase().replace(/\s+/g,'_');
    byStatus[s] = (byStatus[s] || 0) + 1;
  });

  var tats      = cases.map(function(c){ return c._tatDays; });
  var validTATs = tats.filter(function(t){ return t !== null && t >= 0; });

  // ── Pipeline funnel counts ────────────────────────
  var lvlCounts = {
    registered : total,
    inReview   : cases.filter(function(c){ return !!c._reviewDate; }).length,
    l1Done     : cases.filter(function(c){ return c._hasL1; }).length,
    l2Done     : cases.filter(function(c){ return c._hasL2; }).length,
    l3Done     : cases.filter(function(c){ return c._hasL3; }).length,
    l4Done     : cases.filter(function(c){ return c._hasL4; }).length,
    onboarded  : cases.filter(function(c){ return !!c._onboardedDate; }).length
  };

  // ── Active cases (not terminal) ───────────────────
  var activeCases = cases.filter(function(c){
    var cs = String(c._currentStatus||'').toUpperCase().replace(/\s+/g,'_');
    return ['ONBOARDED','CHURNED','DEACTIVATED','REACTIVATED'].indexOf(cs) === -1;
  });

  var pendingL1 = activeCases.filter(function(c){ return !c._hasL1 && c._reviewDate; }).length;
  var pendingL2 = activeCases.filter(function(c){ return c._hasL1 && !c._hasL2; }).length;
  var pendingL3 = activeCases.filter(function(c){ return c._hasL2 && !c._hasL3; }).length;

  // ── SLA counts ────────────────────────────────────
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
    var cat = c._category || 'Other';
    if (!catMap[cat]) catMap[cat] = { total:0, onboarded:0, churned:0, deactivated:0, tats:[], ltv:0, active:0 };
    catMap[cat].total++;
    var cs = String(c._currentStatus||'').toUpperCase().replace(/\s+/g,'_');
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

  // ── Vertical breakdown ────────────────────────────
  var vertMap = {};
  cases.forEach(function(c){
    var vt = c._businessVertical || 'Other';
    if (!vertMap[vt]) vertMap[vt] = { total:0, onboarded:0, churned:0, deactivated:0, active:0, tats:[], ltv:0 };
    vertMap[vt].total++;
    var cs = String(c._currentStatus||'').toUpperCase().replace(/\s+/g,'_');
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
    var st = c._state || 'Unknown';
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
    if (!c._createdDate) return;
    var d = parseDate(c._createdDate);
    if (!d || isNaN(d)) return;
    var key = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
    if (!monthlyMap[key]) monthlyMap[key] = { total:0, onboarded:0 };
    monthlyMap[key].total++;
    if (c._onboardedDate) monthlyMap[key].onboarded++;
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
      return { id:c._id, name:c._name, category:c._category, ltv:c._ltv, status:c._currentStatus };
    });

  // ── Pending action queue ──────────────────────────
  var pendingActions = {
    awaitingDocs : activeCases.filter(function(c){ return !c._reviewDate; }).length,
    awaitingL1   : pendingL1,
    awaitingL2   : pendingL2,
    awaitingL3   : pendingL3,
    sentBack     : cases.filter(function(c){
                     var cs = String(c._currentStatus||'').toUpperCase().replace(/\s+/g,'_');
                     return cs === 'SENT_BACK' || cs === 'SENTBACK';
                   }).length,
    tatBreached  : slaMap.overdue
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

  // ── Cases for table (ALL rows — no arbitrary cap) ─
  var tableCases = cases.map(function(c){
    return {
      id               : c._id,
      name             : c._name,
      gstin            : c._gstin,
      city             : c._city,
      state            : c._state,
      category         : c._category,
      businessVertical : c._businessVertical,
      vendorType       : c._vendorType,
      partyType        : c._partyType,
      status           : String(c.status||'').trim(),
      currentStatus    : c._currentStatus,
      stage            : c._stage,
      createdDate      : c._createdDate,
      onboardedDate    : c._onboardedDate,
      reviewDate       : c._reviewDate,
      statusChangedOn  : c._statusChangedOn,
      tatDays          : c._tatDays,
      ageDays          : c._agedays,
      stageDays        : c._stageDays,
      slaStatus        : c._slaStatus,
      level1           : c._level1,
      level2           : c._level2,
      level3           : c._level3,
      level4           : c._level4,
      ltv              : c._ltv,
      firstShipment    : c._firstShipment,
      l1r1             : c._l1r1,
      l1r2             : c._l1r2,
      l2r1             : c._l2r1,
      l2r2             : c._l2r2,
      l3r1             : c._l3r1,
      l3r2             : c._l3r2,
      l4r1             : c._l4r1,
      l4r2             : c._l4r2,
      contactPerson    : c._contact,
      remarks          : c._remarks
    };
  });

  return {
    generatedAt    : new Date().toLocaleString('en-IN'),
    total          : total,
    byStatus       : byStatus,
    avgTAT         : avg(validTATs),
    sla            : slaMap,
    pipeline       : lvlCounts,
    pendingActions : pendingActions,
    stageBreakdown : stageBreakdown,
    categories     : categories,
    verticals      : verticals,
    states         : states,
    tatBuckets     : tatBuckets,
    monthlyTrend   : monthlyTrend,
    topLTV         : topLTV,
    cases          : tableCases,
    // Debug: actual column headers detected in the sheet
    _sheetHeaders  : hdrs
  };
}
