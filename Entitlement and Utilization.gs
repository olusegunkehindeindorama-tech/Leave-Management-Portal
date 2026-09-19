/**
 * ============================================================
 *  ENTITLEMENT & UTILIZATION  (policy-driven, no hard-coded multipliers)
 * ============================================================
 *
 *  Rules (verbatim):
 *  1. Sys_LeavePolicies is authoritative (calc method, deduct-from,
 *     carry-forward deadline, Multiplier Yes/No, entitlements, lifecycle).
 *  2. Multiplier column on the matched policy = "Yes" | "No".
 *     Multiplier_Policy sheet supplies the numeric weight for each shift
 *     code under the Yes column or the No column accordingly.
 *  3. Previous-year balance (StartingBal) may only be used when the leave
 *     START date is on or before the Carry Forward Deadline (month/day
 *     applied to the current calendar year).
 *  4. tblShift is WIDE (Emp ID | yyyy-MM-dd | …). Empty cell → default
 *     Mon–Fri = "G", Sat–Sun = "O".
 *  5. Master data loaded once per batch run (policies, multipliers,
 *     employees, shifts, leaves, starting bal).
 *  6. calculateLeaveUtilized / recalculateAllLeaveUtilized = on-click.
 *     calculateLeaveUtilize = fast path for UI submit.
 *  7. If a leave spans the carry-forward deadline it is SPLIT into two
 *     rows (pre-deadline → prev entitlement year, post → current).
 *     Wrong splits are merged then re-split correctly.
 *  8. generateYearEndStartingBalances + Jan 1 trigger write
 *     "{year} Balance" into StartingBal for the year that just ended.
 * ============================================================
 */

// ---------------------------------------------------------------------------
// Policy matching helpers
// ---------------------------------------------------------------------------

function checkMatch(empValue, policyValue, weight) {
  if (String(policyValue || '').trim().toUpperCase() === 'DEFAULT') return 0;
  if (String(empValue || '').toLowerCase() === String(policyValue || '').toLowerCase()) return weight;
  return -1;
}

function mapCategoryToInitials(fullCategory) {
  var cat = String(fullCategory || '').toLowerCase();
  if (cat.indexOf('junior staff') > -1) return 'JS';
  if (cat.indexOf('non union mgt') > -1) return 'NUMS';
  if (cat.indexOf('non union senior') > -1) return 'NUSS';
  if (cat.indexOf('senior staff') > -1) return 'SS';
  if (cat.indexOf('trainee') > -1) return 'T';
  return fullCategory;
}

function parseDDMMYYYY(dateString) {
  if (!dateString && dateString !== 0) return new Date();
  if (dateString instanceof Date) return dateString;
  var s = String(dateString).trim();
  var parts = s.split('/');
  if (parts.length === 3) return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? new Date() : d;
}

function evaluateLifecycle(doj, ruleString, todayDate) {
  if (!ruleString || ruleString === 'DEFAULT' || String(ruleString).trim() === '') return true;
  if (!(doj instanceof Date) || isNaN(doj.getTime())) return true;
  var regex = /([><=]+)\s*(\d+)\s*(year|years|month|months|day|days)/i;
  var match = String(ruleString).match(regex);
  if (!match) return true;
  var operator = match[1];
  var value = parseFloat(match[2]);
  var unit = match[3].toLowerCase();
  var diffDays = (todayDate.getTime() - doj.getTime()) / (1000 * 3600 * 24);
  var targetDays = 0;
  if (unit.indexOf('year') > -1) targetDays = value * 365.25;
  else if (unit.indexOf('month') > -1) targetDays = value * 30.4375;
  else targetDays = value;
  switch (operator) {
    case '>=': return diffDays >= targetDays;
    case '>': return diffDays > targetDays;
    case '<=': return diffDays <= targetDays;
    case '<': return diffDays < targetDays;
    case '==':
    case '=': return Math.round(diffDays) === Math.round(targetDays);
    default: return true;
  }
}

function formatDateKey(dateObj) {
  var y = dateObj.getFullYear();
  var m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  var d = ('0' + dateObj.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

function defaultShiftCode_(dateObj) {
  var day = dateObj.getDay();
  return (day === 0 || day === 6) ? 'O' : 'G';
}

// ---------------------------------------------------------------------------
// Master data loader (one shot for batch jobs)
// ---------------------------------------------------------------------------

/**
 * Load every sheet needed for entitlement / utilization into memory.
 * @returns {Object} master
 */
function loadEntitlementMasterData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = {
    policies: [],
    multipliers: { Yes: {}, No: {} }, // shiftCode → weight
    employees: {},
    shifts: {},      // empId → { 'yyyy-MM-dd': code }
    leaves: [],      // raw leave row objects + _row index
    leaveHeaders: [],
    startingBal: {}, // empId → { '2025 Balance': n, ... }
    startingBalHeaders: [],
    sheet: {
      leave: ss.getSheetByName('tblLeave'),
      shift: ss.getSheetByName('tblShift'),
      emp: ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee'),
      policy: ss.getSheetByName('Sys_LeavePolicies'),
      mult: ss.getSheetByName('Multiplier_Policy') || ss.getSheetByName('Multiplier Policy'),
      startBal: ss.getSheetByName('StartingBal')
    }
  };

  // --- Policies ---
  if (master.sheet.policy) {
    var pData = master.sheet.policy.getDataRange().getValues();
    if (pData.length > 1) {
      var pH = pData[0].map(function (h) { return String(h).trim(); });
      for (var i = 1; i < pData.length; i++) {
        var o = {};
        for (var c = 0; c < pH.length; c++) o[pH[c]] = pData[i][c];
        master.policies.push(o);
      }
    }
  }

  // --- Multiplier_Policy ---
  // Expected: col0 = Shift/Code, then columns named Yes / No (or Multiplier Yes / Multiplier No)
  if (master.sheet.mult) {
    var mData = master.sheet.mult.getDataRange().getValues();
    if (mData.length > 1) {
      var mH = mData[0].map(function (h) { return String(h).trim(); });
      var yesIdx = -1, noIdx = -1, codeIdx = 0;
      for (var mh = 0; mh < mH.length; mh++) {
        var hn = mH[mh].toLowerCase();
        if (hn === 'yes' || hn === 'multiplier yes' || hn === 'multiplier_yes') yesIdx = mh;
        else if (hn === 'no' || hn === 'multiplier no' || hn === 'multiplier_no') noIdx = mh;
        else if (hn === 'shift' || hn === 'code' || hn === 'shift code' || hn === 'shiftcode') codeIdx = mh;
      }
      // If Yes/No not found by name, assume col1=Yes col2=No
      if (yesIdx < 0 && mH.length > 1) yesIdx = 1;
      if (noIdx < 0 && mH.length > 2) noIdx = 2;

      for (var mr = 1; mr < mData.length; mr++) {
        var code = String(mData[mr][codeIdx] || '').trim().toUpperCase();
        if (!code) continue;
        if (yesIdx >= 0) master.multipliers.Yes[code] = Number(mData[mr][yesIdx]);
        if (noIdx >= 0) master.multipliers.No[code] = Number(mData[mr][noIdx]);
      }
    }
  }

  // --- Employees ---
  if (master.sheet.emp) {
    var eData = master.sheet.emp.getDataRange().getValues();
    if (eData.length > 1) {
      var eH = eData[0].map(function (h) { return String(h).trim(); });
      var idIdx = eH.indexOf('Emp ID');
      if (idIdx < 0) idIdx = 1;
      for (var er = 1; er < eData.length; er++) {
        var eid = String(eData[er][idIdx] || '').trim().toUpperCase();
        if (!eid) continue;
        var eo = {};
        for (var ec = 0; ec < eH.length; ec++) eo[eH[ec]] = eData[er][ec];
        master.employees[eid] = eo;
      }
    }
  }

  // --- Shifts (WIDE) ---
  if (master.sheet.shift) {
    var sData = master.sheet.shift.getDataRange().getValues();
    if (sData.length > 1) {
      var h1 = String(sData[0][1] || '').trim();
      var isWide = /^\d{4}-\d{2}-\d{2}/.test(h1) ||
        (String(sData[0][0] || '').toLowerCase().indexOf('emp') === 0 && h1.toLowerCase() !== 'date');

      if (isWide) {
        var dateHeaders = [];
        for (var sc = 1; sc < sData[0].length; sc++) {
          var ds = String(sData[0][sc] || '').trim();
          if (ds.length >= 10) ds = ds.substring(0, 10);
          dateHeaders.push(ds);
        }
        for (var sr = 1; sr < sData.length; sr++) {
          var sid = String(sData[sr][0] || '').trim().toUpperCase();
          if (!sid) continue;
          if (!master.shifts[sid]) master.shifts[sid] = {};
          for (var sc2 = 1; sc2 < sData[sr].length; sc2++) {
            var code2 = String(sData[sr][sc2] || '').trim().toUpperCase();
            if (code2 && dateHeaders[sc2 - 1]) master.shifts[sid][dateHeaders[sc2 - 1]] = code2;
          }
        }
      } else {
        // legacy long
        for (var lr = 1; lr < sData.length; lr++) {
          var lid = String(sData[lr][0] || '').trim().toUpperCase();
          if (!lid) continue;
          var dVal = sData[lr][1];
          var dk;
          if (dVal instanceof Date) dk = formatDateKey(dVal);
          else dk = String(dVal || '').trim().substring(0, 10);
          if (!dk) continue;
          if (!master.shifts[lid]) master.shifts[lid] = {};
          master.shifts[lid][dk] = String(sData[lr][2] || '').trim().toUpperCase();
        }
      }
    }
  }

  // --- Leaves ---
  if (master.sheet.leave) {
    var lData = master.sheet.leave.getDataRange().getValues();
    if (lData.length > 0) {
      master.leaveHeaders = lData[0].map(function (h) { return String(h).trim(); });
      for (var li = 1; li < lData.length; li++) {
        var lo = { _row: li + 1 };
        for (var lc = 0; lc < master.leaveHeaders.length; lc++) {
          lo[master.leaveHeaders[lc]] = lData[li][lc];
        }
        if (String(lo['Emp ID'] || '').trim()) master.leaves.push(lo);
      }
    }
  }

  // --- StartingBal ---
  if (master.sheet.startBal) {
    var bData = master.sheet.startBal.getDataRange().getValues();
    if (bData.length > 0) {
      master.startingBalHeaders = bData[0].map(function (h) { return String(h).trim(); });
      var beIdx = master.startingBalHeaders.indexOf('Emp No');
      if (beIdx < 0) beIdx = master.startingBalHeaders.indexOf('Emp ID');
      for (var bi = 1; bi < bData.length; bi++) {
        var bid = beIdx >= 0 ? String(bData[bi][beIdx] || '').trim().toUpperCase() : '';
        if (!bid) continue;
        var bo = {};
        for (var bc = 0; bc < master.startingBalHeaders.length; bc++) {
          bo[master.startingBalHeaders[bc]] = bData[bi][bc];
        }
        master.startingBal[bid] = bo;
      }
    }
  }

  return master;
}

// ---------------------------------------------------------------------------
// Policy lookup for a leave type + employee context
// ---------------------------------------------------------------------------

/**
 * Best-matching Sys_LeavePolicies row for this emp + leave type.
 */
function matchPolicyForEmp_(policies, emp, leaveType, today) {
  var bu = String(emp['Business Unit'] || emp.bu || '').trim();
  var catFull = String(emp['Category'] || emp.category || '').trim();
  var catInitials = mapCategoryToInitials(catFull);
  var empStatus = String(emp['Status'] || emp.status || '').trim();
  var gender = String(emp['Gender'] || emp.gender || '').trim();
  var dojRaw = emp['Date of Join'] || emp.doj;
  var doj = (dojRaw instanceof Date) ? dojRaw : parseDDMMYYYY(String(dojRaw || ''));
  today = today || new Date();

  var best = null;
  var bestScore = -1;

  for (var i = 0; i < policies.length; i++) {
    var pol = policies[i];
    if (String(pol['Leave Type'] || '').trim() !== String(leaveType || '').trim()) continue;

    var scoreBU = checkMatch(bu.toUpperCase(), String(pol['Business Unit'] || '').trim().toUpperCase(), 100);
    var scoreCat = checkMatch(catInitials, String(pol['Category'] || '').trim(), 10);
    if (scoreCat === -1) scoreCat = checkMatch(catFull, String(pol['Category'] || '').trim(), 10);
    var scoreStatus = checkMatch(empStatus, String(pol['Status'] || '').trim(), 5);
    var scoreGender = checkMatch(gender, String(pol['Gender'] || '').trim(), 1);
    if (scoreBU === -1 || scoreCat === -1 || scoreStatus === -1 || scoreGender === -1) continue;
    if (!evaluateLifecycle(doj, String(pol['Who is entitled (Lifecycle)'] || '').trim(), today)) continue;

    var total = scoreBU + scoreCat + scoreStatus + scoreGender;
    if (total > bestScore) {
      bestScore = total;
      best = pol;
    }
  }
  return best;
}

/**
 * Resolve calc meta from a policy row (or leave type alone).
 */
function policyCalcMeta_(pol) {
  var method = 'ActualDays';
  var multFlag = 'No';
  var deductFrom = '';
  var deadline = null;

  if (pol) {
    var m = String(pol['Calculation Method'] || '').trim();
    var ml = m.toLowerCase().replace(/\s+/g, '');
    if (ml.indexOf('shift') > -1 || ml.indexOf('roaster') > -1 || ml.indexOf('roster') > -1) {
      method = 'ShiftRoaster';
    }
    multFlag = String(pol['Multiplier'] || 'No').trim();
    if (/^y/i.test(multFlag)) multFlag = 'Yes';
    else multFlag = 'No';
    deductFrom = String(pol['Deduct from'] || '').trim();

    var deadlineRaw = pol['Carry Forward Deadline'];
    if (deadlineRaw !== null && deadlineRaw !== undefined && String(deadlineRaw).trim() !== '' &&
        String(deadlineRaw).trim().toLowerCase() !== 'no') {
      if (deadlineRaw instanceof Date && !isNaN(deadlineRaw.getTime())) {
        deadline = new Date(deadlineRaw.getFullYear(), deadlineRaw.getMonth(), deadlineRaw.getDate());
      } else {
        deadline = parseDDMMYYYY(String(deadlineRaw));
        if (isNaN(deadline.getTime())) deadline = null;
      }
    }
  }
  return { method: method, multFlag: multFlag, deductFrom: deductFrom, deadline: deadline, policy: pol };
}

/**
 * Weight for one shift code from Multiplier_Policy (Yes or No column).
 * Missing code → 1 for working, but we still honour table only — if absent return 1.
 */
function multiplierWeight_(multipliers, multFlag, shiftCode) {
  var table = (multFlag === 'Yes') ? (multipliers.Yes || {}) : (multipliers.No || {});
  var code = String(shiftCode || '').trim().toUpperCase();
  if (Object.prototype.hasOwnProperty.call(table, code) && !isNaN(Number(table[code]))) {
    return Number(table[code]);
  }
  // Code not in table — try common aliases only if exact key missing
  if (code === 'OFF' || code === 'REST' || code === 'WO' || code === 'W/O') {
    if (Object.prototype.hasOwnProperty.call(table, 'O')) return Number(table['O']);
  }
  return 1; // safe neutral if policy table has no entry
}

// ---------------------------------------------------------------------------
// Core utilization (single range)
// ---------------------------------------------------------------------------

/**
 * Fast UI path — uses cache helpers when master not supplied.
 */
function calculateLeaveUtilize(empId, startDate, endDate, leaveType) {
  var sDate = new Date(startDate);
  var eDate = new Date(endDate);
  if (isNaN(sDate.getTime()) || isNaN(eDate.getTime()) || eDate < sDate) return 0;

  var policies = (typeof loadPoliciesCached_ === 'function') ? loadPoliciesCached_() : [];
  var empMap = (typeof loadEmployeeMapCached_ === 'function') ? loadEmployeeMapCached_() : {};
  var emp = empMap[String(empId).trim().toUpperCase()] || { 'Emp ID': empId };
  var pol = matchPolicyForEmp_(policies, emp, leaveType, new Date());
  var meta = policyCalcMeta_(pol);

  var inclusive = Math.round((eDate.getTime() - sDate.getTime()) / 86400000) + 1;
  if (meta.method === 'ActualDays') return inclusive;

  // Multipliers from sheet
  var multipliers = loadMultipliersCached_();
  var shiftMap = (typeof loadShiftMapForEmp_ === 'function') ? loadShiftMapForEmp_(empId) : {};

  var total = 0;
  var curr = new Date(sDate.getFullYear(), sDate.getMonth(), sDate.getDate());
  var end = new Date(eDate.getFullYear(), eDate.getMonth(), eDate.getDate());
  while (curr <= end) {
    var key = formatDateKey(curr);
    var code = shiftMap[key];
    if (!code) code = defaultShiftCode_(curr);
    total += multiplierWeight_(multipliers, meta.multFlag, code);
    curr.setDate(curr.getDate() + 1);
  }
  return total;
}

/** Same as above but with preloaded master (batch). */
function calculateLeaveUtilizeWithMaster_(master, empId, startDate, endDate, leaveType) {
  var sDate = new Date(startDate);
  var eDate = new Date(endDate);
  if (isNaN(sDate.getTime()) || isNaN(eDate.getTime()) || eDate < sDate) return 0;

  var emp = master.employees[String(empId).trim().toUpperCase()] || { 'Emp ID': empId };
  var pol = matchPolicyForEmp_(master.policies, emp, leaveType, new Date());
  var meta = policyCalcMeta_(pol);

  var inclusive = Math.round((eDate.getTime() - sDate.getTime()) / 86400000) + 1;
  if (meta.method === 'ActualDays') return inclusive;

  var shiftMap = master.shifts[String(empId).trim().toUpperCase()] || {};
  var total = 0;
  var curr = new Date(sDate.getFullYear(), sDate.getMonth(), sDate.getDate());
  var end = new Date(eDate.getFullYear(), eDate.getMonth(), eDate.getDate());
  while (curr <= end) {
    var key = formatDateKey(curr);
    var code = shiftMap[key];
    if (!code) code = defaultShiftCode_(curr);
    total += multiplierWeight_(master.multipliers, meta.multFlag, code);
    curr.setDate(curr.getDate() + 1);
  }
  return total;
}

/** Cached Multiplier_Policy (short TTL). */
function loadMultipliersCached_() {
  if (typeof cacheGet_ === 'function') {
    var hit = cacheGet_('mult_pol');
    if (hit) return hit;
  }
  var master = { multipliers: { Yes: {}, No: {} } };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Multiplier_Policy') || ss.getSheetByName('Multiplier Policy');
  if (sh) {
    var mData = sh.getDataRange().getValues();
    if (mData.length > 1) {
      var mH = mData[0].map(function (h) { return String(h).trim(); });
      var yesIdx = -1, noIdx = -1, codeIdx = 0;
      for (var mh = 0; mh < mH.length; mh++) {
        var hn = mH[mh].toLowerCase();
        if (hn === 'yes' || hn === 'multiplier yes') yesIdx = mh;
        else if (hn === 'no' || hn === 'multiplier no') noIdx = mh;
        else if (hn === 'shift' || hn === 'code' || hn === 'shift code') codeIdx = mh;
      }
      if (yesIdx < 0 && mH.length > 1) yesIdx = 1;
      if (noIdx < 0 && mH.length > 2) noIdx = 2;
      for (var mr = 1; mr < mData.length; mr++) {
        var code = String(mData[mr][codeIdx] || '').trim().toUpperCase();
        if (!code) continue;
        if (yesIdx >= 0) master.multipliers.Yes[code] = Number(mData[mr][yesIdx]);
        if (noIdx >= 0) master.multipliers.No[code] = Number(mData[mr][noIdx]);
      }
    }
  }
  if (typeof cachePut_ === 'function') cachePut_('mult_pol', master.multipliers);
  return master.multipliers;
}

// ---------------------------------------------------------------------------
// Carry-forward deadline helpers
// ---------------------------------------------------------------------------

/**
 * Deadline date for a given leave year context.
 * Policy may store a full date (e.g. 2026-06-30); we take month/day and
 * apply to `year` (usually the leave's start year).
 */
function carryDeadlineForYear_(deadlineSrc, year) {
  if (!deadlineSrc) return null;
  var d;
  if (deadlineSrc instanceof Date && !isNaN(deadlineSrc.getTime())) {
    d = new Date(year, deadlineSrc.getMonth(), deadlineSrc.getDate());
  } else {
    var tmp = parseDDMMYYYY(String(deadlineSrc));
    if (isNaN(tmp.getTime())) return null;
    d = new Date(year, tmp.getMonth(), tmp.getDate());
  }
  return d;
}

/**
 * Can this leave start date still draw previous-year balance?
 */
function canUsePrevYearBalance_(startDate, deadlineSrc) {
  if (!deadlineSrc) return false;
  var start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  var deadline = carryDeadlineForYear_(deadlineSrc, start.getFullYear());
  if (!deadline) return false;
  return start.getTime() <= deadline.getTime();
}

// ---------------------------------------------------------------------------
// Split / merge across carry-forward deadline
// ---------------------------------------------------------------------------

/**
 * If leave spans deadline, return two segments; else one.
 * Segment: { start, end, entitlementYear }
 */
function splitLeaveByCarryForward_(startDate, endDate, leaveType, pol) {
  var meta = policyCalcMeta_(pol);
  var s = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  var e = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  var segments = [];

  if (!meta.deadline) {
    segments.push({ start: s, end: e, entitlementYear: s.getFullYear() });
    return segments;
  }

  // Apply deadline to the year of the start date
  var deadline = carryDeadlineForYear_(meta.deadline, s.getFullYear());
  if (!deadline || e.getTime() <= deadline.getTime() || s.getTime() > deadline.getTime()) {
    // Entirely before/on deadline OR entirely after → no split
    var year = s.getFullYear();
    if (s.getTime() > deadline.getTime()) {
      // After deadline — current year only (cannot use prev bal)
      year = s.getFullYear();
    } else {
      // On or before deadline — may use prev bal; charge year = start year
      // Days on/before deadline that consume prev bal still tagged with start year;
      // entitlement year is the leave year of the balance used.
      // Convention: pre-deadline portion → previous calendar year when
      // the leave starts in year Y and deadline is in Y (carry from Y-1).
      year = s.getFullYear();
    }
    segments.push({ start: s, end: e, entitlementYear: year });
    return segments;
  }

  // Spans deadline: split
  var preEnd = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  var postStart = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate() + 1);

  // Pre-deadline: charged against previous-year balance window → Entitlement Year = startYear - 1
  // when carry is from prior year; else startYear. Standard: CF from (Y-1) expires mid Y,
  // so pre-deadline days use Entitlement Year = Y-1, post use Y.
  segments.push({
    start: s,
    end: preEnd,
    entitlementYear: s.getFullYear() - 1
  });
  segments.push({
    start: postStart,
    end: e,
    entitlementYear: postStart.getFullYear()
  });
  return segments;
}

/**
 * Detect adjacent same-emp/type rows that should be one leave (wrong split).
 * Returns groups of leave objects to merge.
 */
function findMergeCandidates_(leaves) {
  // Sort by emp, type, start
  var sorted = leaves.slice().sort(function (a, b) {
    var ae = String(a['Emp ID'] || '').toUpperCase();
    var be = String(b['Emp ID'] || '').toUpperCase();
    if (ae !== be) return ae < be ? -1 : 1;
    var at = String(a['Leave Type'] || '');
    var bt = String(b['Leave Type'] || '');
    if (at !== bt) return at < bt ? -1 : 1;
    return new Date(a['Start Date']) - new Date(b['Start Date']);
  });

  var groups = [];
  var cur = null;

  for (var i = 0; i < sorted.length; i++) {
    var row = sorted[i];
    var s = new Date(row['Start Date']);
    var e = new Date(row['End Date']);
    if (isNaN(s.getTime()) || isNaN(e.getTime())) continue;

    if (!cur) {
      cur = { rows: [row], start: s, end: e };
      continue;
    }

    var sameEmp = String(cur.rows[0]['Emp ID']).toUpperCase() === String(row['Emp ID']).toUpperCase();
    var sameType = String(cur.rows[0]['Leave Type']) === String(row['Leave Type']);
    var nextDay = new Date(cur.end.getFullYear(), cur.end.getMonth(), cur.end.getDate() + 1);
    var adjacent = nextDay.getTime() === new Date(s.getFullYear(), s.getMonth(), s.getDate()).getTime();
    // Only auto-merge if marked as split siblings (same base entry code prefix or reason)
    var reasonA = String(cur.rows[0]['Leave Reason'] || '');
    var reasonB = String(row['Leave Reason'] || '');
    var sameReason = reasonA === reasonB;
    var codeA = String(cur.rows[0]['Entry Code'] || '');
    var codeB = String(row['Entry Code'] || '');
    var splitSibling = (codeA.indexOf('-S1') > -1 || codeA.indexOf('-S2') > -1 ||
      codeB.indexOf('-S1') > -1 || codeB.indexOf('-S2') > -1 ||
      codeA.replace(/-S[12]$/, '') === codeB.replace(/-S[12]$/, ''));

    if (sameEmp && sameType && adjacent && (sameReason || splitSibling)) {
      cur.rows.push(row);
      cur.end = e;
    } else {
      if (cur.rows.length > 1) groups.push(cur);
      cur = { rows: [row], start: s, end: e };
    }
  }
  if (cur && cur.rows.length > 1) groups.push(cur);
  return groups;
}

// ---------------------------------------------------------------------------
// On-click full recalculation
// ---------------------------------------------------------------------------

/**
 * Recalculate No of Days, Leave Utilized, Entitlement Year for every leave.
 * Merges wrong splits, then splits correctly on carry-forward boundary.
 * All policy / multiplier / shift data from sheets (no hard-coded weights).
 */
function calculateLeaveUtilized() {
  var started = new Date().getTime();
  var master = loadEntitlementMasterData_();
  if (!master.sheet.leave) {
    return { success: false, message: 'tblLeave sheet missing.', updated: 0 };
  }
  if (!master.policies.length) {
    return { success: false, message: 'Sys_LeavePolicies empty — cannot recalculate.', updated: 0 };
  }
  if (!Object.keys(master.multipliers.Yes).length && !Object.keys(master.multipliers.No).length) {
    Logger.log('WARNING: Multiplier_Policy sheet missing or empty. Weights default to 1.');
  }

  var headers = master.leaveHeaders;
  var empIdx = headers.indexOf('Emp ID');
  var typeIdx = headers.indexOf('Leave Type');
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var utilIdx = headers.indexOf('Leave Utilized');
  var daysIdx = headers.indexOf('No of Days');
  var yearIdx = headers.indexOf('Entitlement Year');
  var entryIdx = headers.indexOf('Entry Code');
  var reasonIdx = headers.indexOf('Leave Reason');

  // --- 1. Merge wrong splits (in memory) ---
  var mergeGroups = findMergeCandidates_(master.leaves);
  var rowsToDelete = {}; // row number → true
  var mergedLeaves = []; // replacement single leaves

  mergeGroups.forEach(function (g) {
    // Keep the first row's metadata; expand dates
    var base = g.rows[0];
    for (var k = 1; k < g.rows.length; k++) {
      rowsToDelete[g.rows[k]._row] = true;
    }
    base['Start Date'] = g.start;
    base['End Date'] = g.end;
    // Strip -S1/-S2 suffix from entry code after merge
    if (entryIdx >= 0) {
      base['Entry Code'] = String(base['Entry Code'] || '').replace(/-S[12]$/, '');
    }
    mergedLeaves.push(base);
  });

  // Working set = leaves not deleted + merge results already in set
  var working = [];
  master.leaves.forEach(function (lv) {
    if (rowsToDelete[lv._row]) return;
    // if this was a merge base, use updated dates from mergedLeaves
    var replaced = false;
    for (var m = 0; m < mergedLeaves.length; m++) {
      if (mergedLeaves[m]._row === lv._row) {
        working.push(mergedLeaves[m]);
        replaced = true;
        break;
      }
    }
    if (!replaced) working.push(lv);
  });

  // --- 2. Split + compute utilization ---
  var updates = []; // { row, start, end, util, days, year, entryCode? }
  var appends = []; // brand-new rows for post-deadline segment
  var splitCount = 0;
  var mergeCount = mergeGroups.length;

  working.forEach(function (lv) {
    var empId = String(lv['Emp ID'] || '').trim().toUpperCase();
    var lt = String(lv['Leave Type'] || '').trim();
    var s = new Date(lv['Start Date']);
    var e = new Date(lv['End Date']);
    if (!empId || isNaN(s.getTime()) || isNaN(e.getTime())) return;

    var emp = master.employees[empId] || { 'Emp ID': empId };
    var pol = matchPolicyForEmp_(master.policies, emp, lt, new Date());
    var segments = splitLeaveByCarryForward_(s, e, lt, pol);

    if (segments.length === 1) {
      var util = calculateLeaveUtilizeWithMaster_(master, empId, segments[0].start, segments[0].end, lt);
      var days = Math.round((segments[0].end - segments[0].start) / 86400000) + 1;
      updates.push({
        row: lv._row,
        start: segments[0].start,
        end: segments[0].end,
        util: util,
        days: days,
        year: segments[0].entitlementYear,
        entryCode: lv['Entry Code']
      });
    } else {
      // Update first segment in place; append second
      splitCount++;
      var u0 = calculateLeaveUtilizeWithMaster_(master, empId, segments[0].start, segments[0].end, lt);
      var d0 = Math.round((segments[0].end - segments[0].start) / 86400000) + 1;
      var baseCode = String(lv['Entry Code'] || 'LV').replace(/-S[12]$/, '');
      updates.push({
        row: lv._row,
        start: segments[0].start,
        end: segments[0].end,
        util: u0,
        days: d0,
        year: segments[0].entitlementYear,
        entryCode: baseCode + '-S1'
      });

      var u1 = calculateLeaveUtilizeWithMaster_(master, empId, segments[1].start, segments[1].end, lt);
      var d1 = Math.round((segments[1].end - segments[1].start) / 86400000) + 1;
      var newRow = buildBlankLeaveRow_(headers);
      // copy all fields from original
      for (var c = 0; c < headers.length; c++) {
        newRow[c] = lv[headers[c]] !== undefined ? lv[headers[c]] : '';
      }
      if (entryIdx >= 0) newRow[entryIdx] = baseCode + '-S2';
      if (startIdx >= 0) newRow[startIdx] = segments[1].start;
      if (endIdx >= 0) newRow[endIdx] = segments[1].end;
      if (utilIdx >= 0) newRow[utilIdx] = u1;
      if (daysIdx >= 0) newRow[daysIdx] = d1;
      if (yearIdx >= 0) newRow[yearIdx] = segments[1].entitlementYear;
      appends.push(newRow);
    }
  });

  // --- 3. Write updates ---
  var leaveSheet = master.sheet.leave;
  updates.forEach(function (u) {
    if (startIdx >= 0) leaveSheet.getRange(u.row, startIdx + 1).setValue(u.start);
    if (endIdx >= 0) leaveSheet.getRange(u.row, endIdx + 1).setValue(u.end);
    if (utilIdx >= 0) leaveSheet.getRange(u.row, utilIdx + 1).setValue(u.util);
    if (daysIdx >= 0) leaveSheet.getRange(u.row, daysIdx + 1).setValue(u.days);
    if (yearIdx >= 0) leaveSheet.getRange(u.row, yearIdx + 1).setValue(u.year);
    if (entryIdx >= 0 && u.entryCode) leaveSheet.getRange(u.row, entryIdx + 1).setValue(u.entryCode);
  });

  // Delete merged-away rows (bottom-up)
  var delRows = Object.keys(rowsToDelete).map(Number).sort(function (a, b) { return b - a; });
  delRows.forEach(function (rn) {
    leaveSheet.deleteRow(rn);
  });

  if (appends.length) {
    appendLeaveRows_(leaveSheet, appends);
  }

  // Refresh employee balance columns after recalc
  try {
    refreshAllEmployeeLeaveColumns_(master);
  } catch (e) {
    Logger.log('Balance column refresh skipped: ' + e.message);
  }

  if (typeof cacheClearAll_ === 'function') cacheClearAll_();

  var ms = new Date().getTime() - started;
  var msg = 'Recalculated ' + updates.length + ' leave(s), merged ' + mergeCount +
    ' wrong split group(s), created ' + splitCount + ' carry-forward split(s), ' +
    'appended ' + appends.length + ' row(s) in ' + ms + ' ms.';
  Logger.log(msg);
  return {
    success: true,
    message: msg,
    updated: updates.length,
    merged: mergeCount,
    split: splitCount,
    appended: appends.length,
    elapsedMs: ms
  };
}

function recalculateAllLeaveUtilized() {
  return calculateLeaveUtilized();
}

/**
 * After recalc, push balances into tblEmployee leave columns (in-memory batch).
 */
function refreshAllEmployeeLeaveColumns_(master) {
  if (typeof batchComputeEmployeeBalances_ !== 'function') return;
  if (!master.sheet.emp) return;

  var today = new Date();
  var prevYear = today.getFullYear() - 1;
  var leaveUsageByEmp = {};
  master.leaves.forEach(function (lv) {
    // Note: master.leaves may be stale after writes; re-read util from updates is complex.
    // Caller should re-load if needed. Here we use current memory values.
    var id = String(lv['Emp ID'] || '').trim().toUpperCase();
    var lt = String(lv['Leave Type'] || '').trim();
    var util = Number(lv['Leave Utilized']) || 0;
    var ey = Number(lv['Entitlement Year']) || today.getFullYear();
    if (!id || !lt) return;
    if (!leaveUsageByEmp[id]) leaveUsageByEmp[id] = {};
    if (!leaveUsageByEmp[id][lt]) leaveUsageByEmp[id][lt] = {};
    if (!leaveUsageByEmp[id][lt][ey]) leaveUsageByEmp[id][lt][ey] = 0;
    leaveUsageByEmp[id][lt][ey] += util;
  });

  var startingBalMap = {};
  Object.keys(master.startingBal).forEach(function (id) {
    var col = prevYear + ' Balance';
    startingBalMap[id] = Number(master.startingBal[id][col]) || 0;
  });

  var employees = [];
  Object.keys(master.employees).forEach(function (id) {
    employees.push(master.employees[id]);
  });

  batchComputeEmployeeBalances_(employees, master.policies, leaveUsageByEmp, startingBalMap, today, null);

  // Write leave columns back
  var empSheet = master.sheet.emp;
  var eData = empSheet.getDataRange().getValues();
  var eH = eData[0].map(function (h) { return String(h).trim(); });
  var idIdx = eH.indexOf('Emp ID');
  var leaveCols = ['Annual', 'Casual', 'Compassionate', 'Examination', 'Study', 'Maternity', 'Probation'];
  var colIdx = {};
  leaveCols.forEach(function (c) { colIdx[c] = eH.indexOf(c); });

  for (var r = 1; r < eData.length; r++) {
    var eid = String(eData[r][idIdx] || '').trim().toUpperCase();
    var emp = master.employees[eid];
    if (!emp) continue;
    leaveCols.forEach(function (c) {
      if (colIdx[c] >= 0 && emp[c] !== undefined) {
        empSheet.getRange(r + 1, colIdx[c] + 1).setValue(emp[c]);
      }
    });
  }
}

// ---------------------------------------------------------------------------
// Year-end StartingBal generation (Jan 1 trigger)
// ---------------------------------------------------------------------------

/**
 * Write "{yearJustEnded} Balance" column into StartingBal for every employee.
 * Balance = remaining Annual Leave at end of that year (from live calc).
 */
function generateYearEndStartingBalances(optYear) {
  var today = new Date();
  var yearEnded = optYear || (today.getFullYear() - 1);
  var master = loadEntitlementMasterData_();
  if (!master.sheet.startBal) {
    return { success: false, message: 'StartingBal sheet missing.' };
  }

  var colName = yearEnded + ' Balance';
  var headers = master.startingBalHeaders.slice();
  var colIdx = headers.indexOf(colName);
  if (colIdx < 0) {
    // Append new column
    colIdx = headers.length;
    headers.push(colName);
    master.sheet.startBal.getRange(1, colIdx + 1).setValue(colName);
  }

  var empIdIdx = headers.indexOf('Emp No');
  if (empIdIdx < 0) empIdIdx = headers.indexOf('Emp ID');
  if (empIdIdx < 0) empIdIdx = 1;

  // Build usage for yearEnded only
  var leaveUsageByEmp = {};
  master.leaves.forEach(function (lv) {
    var id = String(lv['Emp ID'] || '').trim().toUpperCase();
    var lt = String(lv['Leave Type'] || '').trim();
    var util = Number(lv['Leave Utilized']) || 0;
    var ey = Number(lv['Entitlement Year']) || 0;
    if (!id || !lt || ey !== yearEnded) return;
    if (!leaveUsageByEmp[id]) leaveUsageByEmp[id] = {};
    if (!leaveUsageByEmp[id][lt]) leaveUsageByEmp[id][lt] = {};
    if (!leaveUsageByEmp[id][lt][ey]) leaveUsageByEmp[id][lt][ey] = 0;
    leaveUsageByEmp[id][lt][ey] += util;
  });

  // For year-end snapshot, treat "today" as Dec 31 of yearEnded so CF deadline
  // within that year is applied correctly
  var asOf = new Date(yearEnded, 11, 31);
  var startingBalMap = {}; // prior carry into yearEnded
  var priorCol = (yearEnded - 1) + ' Balance';
  Object.keys(master.startingBal).forEach(function (id) {
    startingBalMap[id] = Number(master.startingBal[id][priorCol]) || 0;
  });

  var employees = [];
  Object.keys(master.employees).forEach(function (id) {
    employees.push(master.employees[id]);
  });

  if (typeof batchComputeEmployeeBalances_ === 'function') {
    batchComputeEmployeeBalances_(employees, master.policies, leaveUsageByEmp, startingBalMap, asOf, null);
  }

  // Ensure every emp has a StartingBal row
  var balSheet = master.sheet.startBal;
  var existing = {};
  var data = balSheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    var id = String(data[r][empIdIdx] || '').trim().toUpperCase();
    if (id) existing[id] = r + 1;
  }

  var written = 0;
  Object.keys(master.employees).forEach(function (id) {
    var emp = master.employees[id];
    var annualBal = emp['Annual'];
    if (annualBal === '' || annualBal === undefined || annualBal === null) annualBal = 0;
    if (annualBal === 'Unlimited') annualBal = 0;

    if (existing[id]) {
      balSheet.getRange(existing[id], colIdx + 1).setValue(Number(annualBal) || 0);
    } else {
      var bu = String(emp['Business Unit'] || '').trim();
      var newRow = [];
      // Build row aligned to headers
      for (var h = 0; h < headers.length; h++) newRow.push('');
      var buIdx = headers.indexOf('BU');
      if (buIdx >= 0) newRow[buIdx] = bu;
      newRow[empIdIdx] = id;
      newRow[colIdx] = Number(annualBal) || 0;
      balSheet.appendRow(newRow);
      existing[id] = balSheet.getLastRow();
    }
    written++;
  });

  return {
    success: true,
    message: 'StartingBal column "' + colName + '" updated for ' + written + ' employee(s).',
    column: colName,
    written: written
  };
}

/** Install time-driven trigger: 1 Jan ~00:00 every year. */
function setupYearEndStartingBalTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'generateYearEndStartingBalances') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('generateYearEndStartingBalances')
    .timeBased()
    .onMonthDay(1)
    .inMonth(1)
    .atHour(0)
    .create();
  Logger.log('Year-end trigger set: generateYearEndStartingBalances on 1 Jan ~00:00');
}

function generateEntitlementMatrix() {
  return { success: false, message: 'Use balance report from the web app for live entitlements.' };
}

// Legacy alias used by older getLeaveCalcMeta_ callers
function getLeaveCalcMeta_(leaveType) {
  var policies = (typeof loadPoliciesCached_ === 'function') ? loadPoliciesCached_() : [];
  for (var i = 0; i < policies.length; i++) {
    if (String(policies[i]['Leave Type'] || '').trim() === String(leaveType || '').trim()) {
      return policyCalcMeta_(policies[i]);
    }
  }
  return { method: 'ActualDays', multFlag: 'No', deductFrom: '', deadline: null };
}

/** @deprecated — multipliers now come from Multiplier_Policy */
function shiftRoasterWeight_(code) {
  var mult = loadMultipliersCached_();
  return multiplierWeight_(mult, 'Yes', code);
}
