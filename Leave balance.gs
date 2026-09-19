/**
 * ============================================================
 *  LEAVE BALANCE (policy-driven)
 * ============================================================
 *
 *  Carry-forward (today ≤ Annual CF deadline):
 *    Prev remaining = StartingBal[prevYear] − (prev-year Annual util
 *      + prev-year util of any type with Deduct from = Annual Leave)
 *    This-year Annual bal = Annual entitlement − (curr Annual util
 *      + curr Casual util + curr Examination util + any other Deduct-from-Annual)
 *    tblEmployee Annual = prev remaining + this-year Annual bal
 *
 *  After CF deadline:
 *    Ignore prev year entirely; only current-year math.
 *
 *  Attached leaves (Casual / Examination with Deduct from = Annual):
 *    Reduce BOTH their own balance AND Annual balance.
 *    e.g. 25 Annual / 7 Casual; used 10 Annual + 5 Casual
 *      → Annual bal 10, Casual bal 2
 *
 *  Public:
 *    updateAllEmployeeLeaveBalances()  — write balances into tblEmployee
 *    exportAnnualCasualBalances()      — download Annual + Casual report
 *    apiGetEmployeeBalance(empId)      — UI single-employee
 * ============================================================
 */

function apiGetEmployeeBalance(empId) {
  empId = String(empId || '').trim().toUpperCase();
  if (!empId) return { error: 'Employee Not Found' };

  var pack = loadBalancePackForEmp_(empId);
  if (pack.error) return pack;

  var result = computeBalancesForEmployeeCore_(
    pack.profile,
    pack.policies,
    pack.usage,
    pack.today
  );

  return {
    profile: {
      id: empId,
      name: pack.profile.empName,
      bu: pack.profile.bu,
      dept: pack.profile.department,
      category: pack.profile.category,
      gender: pack.profile.gender,
      doj: pack.profile.doj
    },
    balances: result.balances,
    entitlements: result.entitlements,
    usage: result.usageOut,
    detail: result.detail,
    entitledTypes: result.entitledTypes
  };
}

/** Load profile, policies, usage, carry for one emp. */
function loadBalancePackForEmp_(empId) {
  var empMap = (typeof loadEmployeeMapCached_ === 'function') ? loadEmployeeMapCached_() : {};
  var emp = empMap[empId];
  if (!emp) return { error: 'Employee Not Found' };

  var policies = (typeof loadPoliciesCached_ === 'function') ? loadPoliciesCached_() : [];
  var leaveRows = (typeof loadLeaveRowsForEmp_ === 'function') ? loadLeaveRowsForEmp_(empId) : [];

  var today = new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;

  var carryGross = getStartingBalanceForEmp_(empId, prevYear);

  var usage = {};
  for (var i = 0; i < leaveRows.length; i++) {
    var lr = leaveRows[i];
    var lt = String(lr['Leave Type'] || '').trim();
    var util = Number(lr['Leave Utilized']) || 0;
    var ey = Number(lr['Entitlement Year']);
    if (!ey) {
      var sd = lr['Start Date'];
      if (sd instanceof Date && !isNaN(sd.getTime())) ey = sd.getFullYear();
      else ey = currentYear;
    }
    if (!lt) continue;
    if (!usage[lt]) usage[lt] = {};
    if (!usage[lt][ey]) usage[lt][ey] = 0;
    usage[lt][ey] += util;
  }

  var dojRaw = emp['Date of Join'];
  var doj = (dojRaw instanceof Date) ? dojRaw : parseDDMMYYYY(String(dojRaw || ''));

  return {
    profile: {
      empId: empId,
      bu: String(emp['Business Unit'] || '').trim(),
      category: String(emp['Category'] || '').trim(),
      status: String(emp['Status'] || '').trim(),
      gender: String(emp['Gender'] || '').trim(),
      empName: String(emp['Emp Name'] || '').trim(),
      department: String(emp['Department'] || '').trim(),
      doj: doj,
      carryGross: carryGross
    },
    policies: policies,
    usage: usage,
    today: today
  };
}

function getStartingBalanceForEmp_(empId, year) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sb = ss.getSheetByName('StartingBal');
    if (!sb) return 0;
    var sbData = sb.getDataRange().getValues();
    if (sbData.length < 2) return 0;
    var sbH = sbData[0].map(function (h) { return String(h).trim(); });
    var eIdx = sbH.indexOf('Emp No');
    if (eIdx < 0) eIdx = sbH.indexOf('Emp ID');
    var bIdx = sbH.indexOf(year + ' Balance');
    if (eIdx < 0 || bIdx < 0) return 0;
    empId = String(empId).trim().toUpperCase();
    for (var s = 1; s < sbData.length; s++) {
      if (String(sbData[s][eIdx]).trim().toUpperCase() === empId) {
        return Number(sbData[s][bIdx]) || 0;
      }
    }
  } catch (e) {}
  return 0;
}

/**
 * BATCH helper for EmployeeSync — mutates employee row objects in memory.
 */
function batchComputeEmployeeBalances_(employees, policies, leaveUsageByEmp, startingBalMap, today, leaveColMap) {
  leaveColMap = leaveColMap || defaultLeaveColMap_();
  today = today || new Date();

  for (var i = 0; i < employees.length; i++) {
    var emp = employees[i];
    var empId = String(emp['Emp ID'] || '').trim().toUpperCase();
    var dojRaw = emp['Date of Join'];
    var doj = (dojRaw instanceof Date) ? dojRaw : parseDDMMYYYY(String(dojRaw || ''));

    var core = computeBalancesForEmployeeCore_(
      {
        empId: empId,
        bu: String(emp['Business Unit'] || '').trim(),
        category: String(emp['Category'] || '').trim(),
        status: String(emp['Status'] || '').trim(),
        gender: String(emp['Gender'] || '').trim(),
        empName: String(emp['Emp Name'] || '').trim(),
        department: String(emp['Department'] || '').trim(),
        doj: doj,
        carryGross: Number(startingBalMap[empId]) || 0
      },
      policies,
      leaveUsageByEmp[empId] || {},
      today
    );

    Object.keys(leaveColMap).forEach(function (col) {
      var leaveType = leaveColMap[col];
      var bal = core.balances[leaveType];
      if (bal === undefined || bal === null) emp[col] = '';
      else if (bal === 'Unlimited') emp[col] = 'Unlimited';
      else emp[col] = Number(bal);
    });
  }
}

function defaultLeaveColMap_() {
  return {
    'Annual': 'Annual Leave',
    'Casual': 'Casual Leave',
    'Compassionate': 'Compassionate Leave',
    'Examination': 'Examination Leave',
    'Study': 'Study Leave',
    'Maternity': 'Maternity Leave',
    'Probation': 'Probation Leave'
  };
}

/**
 * Core pure calculation.
 *
 * usage structure: { 'Annual Leave': { 2025: 7, 2026: 10 }, 'Casual Leave': { 2025: 5, 2026: 5 }, ... }
 */
function computeBalancesForEmployeeCore_(profile, policies, usage, today) {
  today = today || new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;

  var bu = profile.bu;
  var catFull = profile.category;
  var empStatus = profile.status;
  var gender = profile.gender;
  var doj = profile.doj instanceof Date ? profile.doj : parseDDMMYYYY(String(profile.doj || ''));
  var catInitials = mapCategoryToInitials(catFull);
  var carryGross = Number(profile.carryGross) || 0;

  // ---- Match best policy per leave type ----
  var typeMeta = {};
  for (var p = 0; p < policies.length; p++) {
    var pol = policies[p];
    var lType = String(pol['Leave Type'] || '').trim();
    if (!lType) continue;

    var scoreBU = checkMatch(bu.toUpperCase(), String(pol['Business Unit'] || '').trim().toUpperCase(), 100);
    var scoreCat = checkMatch(catInitials, String(pol['Category'] || '').trim(), 10);
    if (scoreCat === -1) scoreCat = checkMatch(catFull, String(pol['Category'] || '').trim(), 10);
    var scoreStatus = checkMatch(empStatus, String(pol['Status'] || '').trim(), 5);
    var scoreGender = checkMatch(gender, String(pol['Gender'] || '').trim(), 1);
    if (scoreBU === -1 || scoreCat === -1 || scoreStatus === -1 || scoreGender === -1) continue;
    if (!evaluateLifecycle(doj, String(pol['Who is entitled (Lifecycle)'] || '').trim(), today)) continue;

    var totalScore = scoreBU + scoreCat + scoreStatus + scoreGender;
    var entRaw = String(pol['Annual Entitlements'] || '').trim();
    var isUnlimited = (entRaw === '' || entRaw.toLowerCase() === 'unlimited');
    var entVal = isUnlimited ? 'Unlimited' : (Number(entRaw) || 0);

    var show = String(pol['Balance Page Show'] || '').trim().toLowerCase();
    var showYes = (show === 'yes' || show === 'y' || show === 'true');
    var deductFrom = String(pol['Deduct from'] || '').trim();
    if (deductFrom.toLowerCase() === 'none' || deductFrom.toLowerCase() === 'null') deductFrom = '';

    var deadlineRaw = pol['Carry Forward Deadline'];
    var deadline = null;
    if (deadlineRaw !== null && deadlineRaw !== undefined &&
        String(deadlineRaw).trim() !== '' && String(deadlineRaw).trim().toLowerCase() !== 'no') {
      if (deadlineRaw instanceof Date && !isNaN(deadlineRaw.getTime())) {
        deadline = new Date(currentYear, deadlineRaw.getMonth(), deadlineRaw.getDate());
      } else {
        var tmp = parseDDMMYYYY(String(deadlineRaw));
        if (!isNaN(tmp.getTime())) deadline = new Date(currentYear, tmp.getMonth(), tmp.getDate());
      }
    }

    if (!typeMeta[lType] || totalScore > typeMeta[lType].score) {
      typeMeta[lType] = {
        score: totalScore,
        entitlement: entVal,
        isUnlimited: isUnlimited,
        show: showYes,
        deductFrom: deductFrom,
        deadline: deadline
      };
    }
  }

  // Types that deduct from Annual (from policy, not hard-coded names)
  var attachedToAnnual = [];
  Object.keys(typeMeta).forEach(function (t) {
    if (typeMeta[t].deductFrom === 'Annual Leave') attachedToAnnual.push(t);
  });

  function utilOf(leaveType, year) {
    var u = usage[leaveType] || {};
    return Number(u[year]) || 0;
  }

  // Composite usage charged against Annual for a given year
  function annualChargeForYear_(year) {
    var total = utilOf('Annual Leave', year);
    for (var a = 0; a < attachedToAnnual.length; a++) {
      total += utilOf(attachedToAnnual[a], year);
    }
    return total;
  }

  // CF window from Annual policy deadline (best matched Annual Leave)
  var annualMeta = typeMeta['Annual Leave'] || null;
  var carryActive = false;
  var carryDeadline = null;
  if (annualMeta && annualMeta.deadline) {
    carryDeadline = annualMeta.deadline;
    carryActive = today.getTime() <= carryDeadline.getTime();
  }

  // Prev-year remaining against StartingBal (only while CF active)
  var prevYearCharge = annualChargeForYear_(prevYear);
  var prevAvail = 0;
  if (carryActive) {
    prevAvail = Math.max(0, carryGross - prevYearCharge);
  }

  var balances = {};
  var entitlements = {};
  var usageOut = {};
  var detail = {};

  Object.keys(typeMeta).forEach(function (t) {
    var meta = typeMeta[t];
    entitlements[t] = meta.entitlement;
    var cy = utilOf(t, currentYear);
    var py = utilOf(t, prevYear);
    usageOut[t] = cy + (carryActive ? py : 0);

    if (meta.isUnlimited) {
      balances[t] = 'Unlimited';
      detail[t] = {
        prevYearBalance: 0,
        prevYearUtilized: py,
        thisYearEntitlement: 'Unlimited',
        thisYearUtilized: cy,
        thisYearBalance: 'Unlimited',
        currentBalance: 'Unlimited',
        carryExpired: !carryActive,
        carryDeadline: carryDeadline ? leaveDateKeySafe_(carryDeadline) : null,
        deductFrom: meta.deductFrom,
        show: meta.show
      };
      return;
    }

    var ent = Number(meta.entitlement) || 0;
    var thisYearBal = 0;
    var prevBalShown = 0;

    if (t === 'Annual Leave') {
      // This year: entitlement − (annual + all attached)
      var currCharge = annualChargeForYear_(currentYear);
      thisYearBal = Math.max(0, ent - currCharge);
      prevBalShown = prevAvail;
      balances[t] = thisYearBal + prevBalShown;
    } else {
      // Own balance only for this year (attached leaves do not get CF pool)
      thisYearBal = Math.max(0, ent - cy);
      prevBalShown = 0;
      balances[t] = thisYearBal;
    }

    detail[t] = {
      prevYearBalance: prevBalShown,
      prevYearUtilized: py,
      prevYearChargeAgainstAnnual: (t === 'Annual Leave') ? prevYearCharge : null,
      thisYearEntitlement: ent,
      thisYearUtilized: cy,
      thisYearBalance: thisYearBal,
      currentBalance: balances[t],
      carryExpired: !carryActive,
      carryDeadline: carryDeadline ? leaveDateKeySafe_(carryDeadline) : null,
      deductFrom: meta.deductFrom,
      show: meta.show
    };
  });

  // Cap attached-leave balances at remaining Annual (cannot take more Casual than Annual allows)
  var annualBal = balances['Annual Leave'];
  if (typeof annualBal === 'number') {
    attachedToAnnual.forEach(function (t) {
      if (typeof balances[t] !== 'number') return;
      var capped = Math.min(balances[t], annualBal);
      balances[t] = capped;
      if (detail[t]) {
        detail[t].currentBalance = capped;
        detail[t].thisYearBalance = capped;
      }
    });
  }

  return {
    balances: balances,
    entitlements: entitlements,
    usageOut: usageOut,
    detail: detail,
    entitledTypes: Object.keys(typeMeta),
    carryActive: carryActive,
    prevAvail: prevAvail
  };
}

function leaveDateKeySafe_(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  try {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return d.getFullYear() + '-' +
      ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
      ('0' + d.getDate()).slice(-2);
  }
}

// ---------------------------------------------------------------------------
// Function 1: Calculate all balances → write tblEmployee columns
// ---------------------------------------------------------------------------

/**
 * Recalculate every employee's leave balances from policies + tblLeave +
 * StartingBal and write into tblEmployee leave columns (Annual, Casual, …).
 */
function updateAllEmployeeLeaveBalances() {
  var started = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  var leaveSheet = ss.getSheetByName('tblLeave');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');
  var sbSheet = ss.getSheetByName('StartingBal');

  if (!empSheet) return { success: false, message: 'tblEmployee missing.' };
  if (!policySheet) return { success: false, message: 'Sys_LeavePolicies missing.' };

  var today = new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;

  // Policies
  var policies = [];
  var pData = policySheet.getDataRange().getValues();
  if (pData.length > 1) {
    var pH = pData[0].map(function (h) { return String(h).trim(); });
    for (var i = 1; i < pData.length; i++) {
      var o = {};
      for (var c = 0; c < pH.length; c++) o[pH[c]] = pData[i][c];
      policies.push(o);
    }
  }

  // StartingBal map
  var startingBalMap = {};
  if (sbSheet) {
    var sbData = sbSheet.getDataRange().getValues();
    if (sbData.length > 1) {
      var sbH = sbData[0].map(function (h) { return String(h).trim(); });
      var eIdx = sbH.indexOf('Emp No');
      if (eIdx < 0) eIdx = sbH.indexOf('Emp ID');
      var bIdx = sbH.indexOf(prevYear + ' Balance');
      if (eIdx >= 0 && bIdx >= 0) {
        for (var s = 1; s < sbData.length; s++) {
          var sid = String(sbData[s][eIdx] || '').trim().toUpperCase();
          if (sid) startingBalMap[sid] = Number(sbData[s][bIdx]) || 0;
        }
      }
    }
  }

  // Leave usage by emp → type → year
  var leaveUsageByEmp = {};
  if (leaveSheet) {
    var lData = leaveSheet.getDataRange().getValues();
    if (lData.length > 1) {
      var lH = lData[0].map(function (h) { return String(h).trim(); });
      var liEmp = lH.indexOf('Emp ID');
      var liType = lH.indexOf('Leave Type');
      var liUtil = lH.indexOf('Leave Utilized');
      var liYear = lH.indexOf('Entitlement Year');
      var liStart = lH.indexOf('Start Date');
      for (var r = 1; r < lData.length; r++) {
        var eid = liEmp >= 0 ? String(lData[r][liEmp] || '').trim().toUpperCase() : '';
        var lt = liType >= 0 ? String(lData[r][liType] || '').trim() : '';
        var util = liUtil >= 0 ? Number(lData[r][liUtil]) || 0 : 0;
        var ey = liYear >= 0 ? Number(lData[r][liYear]) : 0;
        if (!ey && liStart >= 0) {
          var sd = lData[r][liStart];
          if (sd instanceof Date && !isNaN(sd.getTime())) ey = sd.getFullYear();
        }
        if (!eid || !lt) continue;
        if (!leaveUsageByEmp[eid]) leaveUsageByEmp[eid] = {};
        if (!leaveUsageByEmp[eid][lt]) leaveUsageByEmp[eid][lt] = {};
        if (!leaveUsageByEmp[eid][lt][ey]) leaveUsageByEmp[eid][lt][ey] = 0;
        leaveUsageByEmp[eid][lt][ey] += util;
      }
    }
  }

  // Employees
  var eData = empSheet.getDataRange().getValues();
  if (eData.length < 2) return { success: true, message: 'No employees.', updated: 0 };
  var eH = eData[0].map(function (h) { return String(h).trim(); });
  var leaveColMap = defaultLeaveColMap_();
  var colIdx = {};
  Object.keys(leaveColMap).forEach(function (col) {
    colIdx[col] = eH.indexOf(col);
  });
  var idIdx = eH.indexOf('Emp ID');

  var employees = [];
  for (var er = 1; er < eData.length; er++) {
    var rowObj = {};
    for (var ec = 0; ec < eH.length; ec++) rowObj[eH[ec]] = eData[er][ec];
    rowObj._sheetRow = er + 1;
    employees.push(rowObj);
  }

  batchComputeEmployeeBalances_(employees, policies, leaveUsageByEmp, startingBalMap, today, leaveColMap);

  // Write leave columns only
  var updated = 0;
  employees.forEach(function (emp) {
    Object.keys(leaveColMap).forEach(function (col) {
      if (colIdx[col] < 0) return;
      var val = emp[col];
      if (val === undefined) val = '';
      empSheet.getRange(emp._sheetRow, colIdx[col] + 1).setValue(val);
    });
    updated++;
  });

  SpreadsheetApp.flush();
  if (typeof cacheClearAll_ === 'function') cacheClearAll_();

  var ms = new Date().getTime() - started;
  var msg = 'Updated leave balances for ' + updated + ' employee(s) in ' + ms + ' ms.';
  Logger.log(msg);
  return { success: true, message: msg, updated: updated, elapsedMs: ms };
}

// ---------------------------------------------------------------------------
// Function 2: Export Annual + Casual balances for all employees
// ---------------------------------------------------------------------------

/**
 * Build and download a spreadsheet of Annual + Casual balances only.
 * Columns: Emp ID, Emp Name, BU, Department, Category, Gender, DOJ,
 *          Annual Entitlement, Annual Utilized (this year), Annual Balance,
 *          Casual Entitlement, Casual Utilized (this year), Casual Balance,
 *          Prev Year Carry Remaining (if CF active), Carry Deadline
 */
function exportAnnualCasualBalances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!empSheet) {
    return { success: false, message: 'tblEmployee missing.' };
  }

  // Ensure balances are current
  updateAllEmployeeLeaveBalances();

  var today = new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;

  var policies = (typeof loadPoliciesCached_ === 'function') ? loadPoliciesCached_() : [];
  var empMap = (typeof loadEmployeeMapCached_ === 'function') ? loadEmployeeMapCached_() : {};

  // Rebuild usage + starting bal lightly
  var leaveSheet = ss.getSheetByName('tblLeave');
  var leaveUsageByEmp = {};
  if (leaveSheet) {
    var lData = leaveSheet.getDataRange().getValues();
    if (lData.length > 1) {
      var lH = lData[0].map(function (h) { return String(h).trim(); });
      var liEmp = lH.indexOf('Emp ID');
      var liType = lH.indexOf('Leave Type');
      var liUtil = lH.indexOf('Leave Utilized');
      var liYear = lH.indexOf('Entitlement Year');
      var liStart = lH.indexOf('Start Date');
      for (var r = 1; r < lData.length; r++) {
        var eid = liEmp >= 0 ? String(lData[r][liEmp] || '').trim().toUpperCase() : '';
        var lt = liType >= 0 ? String(lData[r][liType] || '').trim() : '';
        var util = liUtil >= 0 ? Number(lData[r][liUtil]) || 0 : 0;
        var ey = liYear >= 0 ? Number(lData[r][liYear]) : 0;
        if (!ey && liStart >= 0) {
          var sd = lData[r][liStart];
          if (sd instanceof Date && !isNaN(sd.getTime())) ey = sd.getFullYear();
        }
        if (!eid || !lt) continue;
        if (!leaveUsageByEmp[eid]) leaveUsageByEmp[eid] = {};
        if (!leaveUsageByEmp[eid][lt]) leaveUsageByEmp[eid][lt] = {};
        if (!leaveUsageByEmp[eid][lt][ey]) leaveUsageByEmp[eid][lt][ey] = 0;
        leaveUsageByEmp[eid][lt][ey] += util;
      }
    }
  }

  var headers = [
    'Emp ID', 'Emp Name', 'BU', 'Department', 'Category', 'Gender', 'Date of Join',
    'Annual Entitlement', 'Annual Utilized (This Year)', 'Annual Balance',
    'Casual Entitlement', 'Casual Utilized (This Year)', 'Casual Balance',
    'Prev Year Carry Remaining', 'Carry Forward Deadline', 'Carry Active'
  ];
  var out = [headers];

  Object.keys(empMap).forEach(function (id) {
    if (id === '_headers') return;
    var emp = empMap[id];
    var dojRaw = emp['Date of Join'];
    var doj = (dojRaw instanceof Date) ? dojRaw : parseDDMMYYYY(String(dojRaw || ''));
    var carryGross = getStartingBalanceForEmp_(id, prevYear);

    var core = computeBalancesForEmployeeCore_(
      {
        empId: id,
        bu: String(emp['Business Unit'] || '').trim(),
        category: String(emp['Category'] || '').trim(),
        status: String(emp['Status'] || '').trim(),
        gender: String(emp['Gender'] || '').trim(),
        empName: String(emp['Emp Name'] || '').trim(),
        department: String(emp['Department'] || '').trim(),
        doj: doj,
        carryGross: carryGross
      },
      policies,
      leaveUsageByEmp[id] || {},
      today
    );

    var dA = core.detail['Annual Leave'] || {};
    var dC = core.detail['Casual Leave'] || {};

    out.push([
      id,
      String(emp['Emp Name'] || ''),
      String(emp['Business Unit'] || ''),
      String(emp['Department'] || ''),
      String(emp['Category'] || ''),
      String(emp['Gender'] || ''),
      doj instanceof Date && !isNaN(doj.getTime()) ? leaveDateKeySafe_(doj) : '',
      dA.thisYearEntitlement != null ? dA.thisYearEntitlement : '',
      dA.thisYearUtilized != null ? dA.thisYearUtilized : 0,
      core.balances['Annual Leave'] != null ? core.balances['Annual Leave'] : '',
      dC.thisYearEntitlement != null ? dC.thisYearEntitlement : '',
      dC.thisYearUtilized != null ? dC.thisYearUtilized : 0,
      core.balances['Casual Leave'] != null ? core.balances['Casual Leave'] : '',
      dA.prevYearBalance != null ? dA.prevYearBalance : 0,
      dA.carryDeadline || '',
      core.carryActive ? 'Yes' : 'No'
    ]);
  });

  // Sort by Emp ID
  var body = out.slice(1).sort(function (a, b) {
    return String(a[0]).localeCompare(String(b[0]));
  });
  out = [headers].concat(body);

  // Create downloadable sheet in Drive (temp) or show as CSV dialog
  var csv = out.map(function (row) {
    return row.map(function (cell) {
      var s = String(cell == null ? '' : cell);
      if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(',');
  }).join('\n');

  var fileName = 'Leave_Balances_Annual_Casual_' +
    Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMdd') + '.csv';

  var encoded = Utilities.base64Encode(Utilities.newBlob(csv, 'text/csv', fileName).getBytes());
  var html =
    '<html><body style="font-family:sans-serif;padding:16px;">' +
    '<h3>Leave Balances Ready</h3>' +
    '<p>' + (out.length - 1) + ' employee(s). Annual + Casual only.</p>' +
    '<a download="' + fileName + '" href="data:text/csv;base64,' + encoded + '" ' +
    'style="display:inline-block;padding:10px 18px;background:#1a73e8;color:#fff;' +
    'text-decoration:none;border-radius:4px;font-weight:bold;" ' +
    'onclick="setTimeout(function(){google.script.host.close();},400);">' +
    'Download CSV</a></body></html>';

  try {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(html).setWidth(360).setHeight(160),
      'Annual & Casual Balances'
    );
  } catch (uiErr) {
    // Running from trigger / no UI — write to Drive instead
    var folderId = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
    try {
      var folder = DriveApp.getFolderById(folderId);
      folder.createFile(Utilities.newBlob(csv, 'text/csv', fileName));
    } catch (de) {
      DriveApp.createFile(Utilities.newBlob(csv, 'text/csv', fileName));
    }
  }

  return {
    success: true,
    message: 'Exported ' + (out.length - 1) + ' employee Annual/Casual balances.',
    rows: out.length - 1
  };
}

// ---------------------------------------------------------------------------
// Legacy report helpers
// ---------------------------------------------------------------------------

function getLeaveBalancesData(optEmpId, optBU, optDept) {
  if (optEmpId) {
    var one = apiGetEmployeeBalance(optEmpId);
    if (one.error) return {};
    var o = {};
    o[optEmpId] = one;
    return o;
  }
  var empMap = loadEmployeeMapCached_();
  var out = {};
  Object.keys(empMap).forEach(function (id) {
    if (id === '_headers') return;
    var emp = empMap[id];
    if (optBU && String(emp['Business Unit'] || '').toUpperCase() !== String(optBU).toUpperCase()) return;
    if (optDept && String(emp['Department'] || '').toUpperCase() !== String(optDept).toUpperCase()) return;
    var bal = apiGetEmployeeBalance(id);
    if (!bal.error) out[id] = bal;
  });
  return out;
}

function generateReportArray(buFilter, deptFilter) {
  var raw = getLeaveBalancesData(null, buFilter, deptFilter);
  var typeList = ['Annual Leave', 'Casual Leave'];
  var headers = ['Emp ID', 'Name', 'Department', 'BU', 'Category'];
  typeList.forEach(function (t) {
    headers.push(t + ' — Prev Year Bal');
    headers.push(t + ' — This Year Entitlement');
    headers.push(t + ' — Utilized This Year');
    headers.push(t + ' — Current Balance');
  });
  var report = [headers];
  Object.keys(raw).forEach(function (id) {
    var emp = raw[id];
    var row = [emp.profile.id, emp.profile.name, emp.profile.dept, emp.profile.bu, emp.profile.category];
    typeList.forEach(function (t) {
      var d = (emp.detail && emp.detail[t]) ? emp.detail[t] : null;
      if (!d) row.push('', '', '', '');
      else row.push(d.prevYearBalance, d.thisYearEntitlement, d.thisYearUtilized, d.currentBalance);
    });
    report.push(row);
  });
  return report;
}
