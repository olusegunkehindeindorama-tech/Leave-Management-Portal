/**
 * Fast leave balance for ONE employee (UI path).
 * Applies Deduct from: Casual + Examination usage reduce Annual balance.
 * Uses Cache.gs helpers when available.
 *
 * Also exports batchComputeEmployeeBalances_ for EmployeeSync.gs (in-memory batch).
 */
function apiGetEmployeeBalance(empId) {
  empId = String(empId || '').trim().toUpperCase();
  if (!empId) return { error: 'Employee Not Found' };

  var empMap = (typeof loadEmployeeMapCached_ === 'function') ? loadEmployeeMapCached_() : {};
  var emp = empMap[empId];
  if (!emp) return { error: 'Employee Not Found' };

  var policies = (typeof loadPoliciesCached_ === 'function') ? loadPoliciesCached_() : [];
  var leaveRows = (typeof loadLeaveRowsForEmp_ === 'function') ? loadLeaveRowsForEmp_(empId) : [];

  var today = new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;

  var bu = String(emp['Business Unit'] || '').trim();
  var catFull = String(emp['Category'] || '').trim();
  var empStatus = String(emp['Status'] || '').trim();
  var gender = String(emp['Gender'] || '').trim();
  var empName = String(emp['Emp Name'] || '').trim();
  var dept = String(emp['Department'] || '').trim();
  var dojRaw = emp['Date of Join'];
  var doj = (dojRaw instanceof Date) ? dojRaw : parseDDMMYYYY(String(dojRaw || ''));
  var catInitials = mapCategoryToInitials(catFull);

  var carryGross = 0;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sb = ss.getSheetByName('StartingBal');
    if (sb) {
      var sbData = sb.getDataRange().getValues();
      var sbH = sbData[0].map(function (h) { return String(h).trim(); });
      var eIdx = sbH.indexOf('Emp No') > -1 ? sbH.indexOf('Emp No') : sbH.indexOf('Emp ID');
      var bIdx = sbH.indexOf(prevYear + ' Balance');
      if (eIdx > -1 && bIdx > -1) {
        for (var s = 1; s < sbData.length; s++) {
          if (String(sbData[s][eIdx]).trim().toUpperCase() === empId) {
            carryGross = Number(sbData[s][bIdx]) || 0;
            break;
          }
        }
      }
    }
  } catch (e) {}

  var usage = {};
  for (var i = 0; i < leaveRows.length; i++) {
    var lr = leaveRows[i];
    var lt = String(lr['Leave Type'] || '').trim();
    var util = Number(lr['Leave Utilized']) || 0;
    var ey = Number(lr['Entitlement Year']) || currentYear;
    if (!lt) continue;
    if (!usage[lt]) usage[lt] = {};
    if (!usage[lt][ey]) usage[lt][ey] = 0;
    usage[lt][ey] += util;
  }

  var result = computeBalancesForEmployeeCore_(
    {
      empId: empId,
      bu: bu,
      category: catFull,
      status: empStatus,
      gender: gender,
      empName: empName,
      department: dept,
      doj: doj,
      carryGross: carryGross
    },
    policies,
    usage,
    today
  );

  return {
    profile: { id: empId, name: empName, bu: bu, dept: dept, category: catFull },
    balances: result.balances,
    entitlements: result.entitlements,
    usage: result.usageOut,
    detail: result.detail,
    entitledTypes: result.entitledTypes
  };
}

/**
 * BATCH: mutates each employee row object in `employees` by writing leave-column balances.
 * Designed for EmployeeSync.gs — zero sheet reads; all inputs preloaded.
 *
 * @param {Object[]} employees  row objects with profile fields (Business Unit, Emp ID, …)
 * @param {Object[]} policies   Sys_LeavePolicies rows as objects
 * @param {Object} leaveUsageByEmp  empId -> leaveType -> year -> utilized
 * @param {Object} startingBalMap   empId -> prev-year carry number
 * @param {Date} today
 * @param {Object} leaveColMap  { 'Annual': 'Annual Leave', ... } sheetCol -> leave type name
 */
function batchComputeEmployeeBalances_(employees, policies, leaveUsageByEmp, startingBalMap, today, leaveColMap) {
  leaveColMap = leaveColMap || {
    'Annual': 'Annual Leave',
    'Casual': 'Casual Leave',
    'Compassionate': 'Compassionate Leave',
    'Examination': 'Examination Leave',
    'Study': 'Study Leave',
    'Maternity': 'Maternity Leave',
    'Probation': 'Probation Leave'
  };

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

    // Write balances into leave-type columns on the employee row
    Object.keys(leaveColMap).forEach(function (col) {
      var leaveType = leaveColMap[col];
      var bal = core.balances[leaveType];
      if (bal === undefined || bal === null) {
        emp[col] = '';
      } else if (bal === 'Unlimited') {
        emp[col] = 'Unlimited';
      } else {
        emp[col] = Number(bal);
      }
    });
  }
}

/**
 * Shared pure calculation (no sheet I/O).
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

  var typeMeta = {};
  for (var p = 0; p < policies.length; p++) {
    var pol = policies[p];
    var lType = String(pol['Leave Type'] || '').trim();
    if (!lType) continue;

    var scoreBU = checkMatch(bu.toUpperCase(), String(pol['Business Unit'] || '').trim().toUpperCase(), 100);
    var scoreCat = checkMatch(catInitials, String(pol['Category'] || '').trim(), 10);
    // Also try full category text against policy category
    if (scoreCat === -1) {
      scoreCat = checkMatch(catFull, String(pol['Category'] || '').trim(), 10);
    }
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
    var deadlineRaw = String(pol['Carry Forward Deadline'] || '').trim();
    var deadline = null;
    if (deadlineRaw && deadlineRaw.toLowerCase() !== 'no') {
      deadline = parseDDMMYYYY(deadlineRaw);
      if (!isNaN(deadline.getTime())) deadline.setFullYear(currentYear);
      else deadline = null;
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

  var balances = {};
  var entitlements = {};
  var usageOut = {};
  var detail = {};

  var annualUsageCurr = 0;
  var annualUsagePrev = 0;
  var casualUsageCurr = 0;
  var examUsageCurr = 0;

  Object.keys(typeMeta).forEach(function (t) {
    var u = usage[t] || {};
    var cy = u[currentYear] || 0;
    var py = u[prevYear] || 0;
    if (t === 'Annual Leave') { annualUsageCurr = cy; annualUsagePrev = py; }
    if (t === 'Casual Leave') casualUsageCurr = cy;
    if (t === 'Examination Leave') examUsageCurr = cy;
  });

  Object.keys(typeMeta).forEach(function (t) {
    var meta = typeMeta[t];
    entitlements[t] = meta.entitlement;
    var u = usage[t] || {};
    var cy = u[currentYear] || 0;
    var py = u[prevYear] || 0;
    usageOut[t] = cy + py;

    var prevAvail = 0;
    var currAvail = 0;
    var total = 0;
    var carryExpired = false;

    if (meta.isUnlimited) {
      total = 'Unlimited';
      currAvail = 'Unlimited';
    } else {
      if (meta.deadline && today > meta.deadline) {
        carryExpired = true;
        prevAvail = 0;
      } else if (t === 'Annual Leave') {
        prevAvail = Math.max(0, carryGross - annualUsagePrev);
      }

      var effectiveCurrUsage = cy;
      if (t === 'Annual Leave') {
        effectiveCurrUsage = annualUsageCurr + casualUsageCurr + examUsageCurr;
      }

      currAvail = Math.max(0, Number(meta.entitlement) - effectiveCurrUsage);
      if (meta.deductFrom === 'Annual Leave') {
        currAvail = Math.max(0, Number(meta.entitlement) - cy);
      }
      total = (typeof currAvail === 'number' ? currAvail : 0) + prevAvail;
    }

    balances[t] = total;
    detail[t] = {
      prevYearBalance: meta.isUnlimited ? 'Unlimited' : prevAvail,
      thisYearEntitlement: meta.entitlement,
      thisYearUtilized: cy,
      thisYearBalance: currAvail,
      currentBalance: total,
      carryExpired: carryExpired,
      deductFrom: meta.deductFrom,
      show: meta.show
    };
  });

  var annualBal = balances['Annual Leave'];
  if (typeof annualBal === 'number') {
    ['Casual Leave', 'Examination Leave'].forEach(function (t) {
      if (!typeMeta[t] || typeMeta[t].deductFrom !== 'Annual Leave') return;
      var own = typeof balances[t] === 'number' ? balances[t] : 0;
      var capped = Math.min(own, annualBal);
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
    entitledTypes: Object.keys(typeMeta)
  };
}

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
  var types = {};
  Object.keys(raw).forEach(function (id) {
    Object.keys(raw[id].detail || {}).forEach(function (t) { types[t] = true; });
  });
  var typeList = Object.keys(types).sort();
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
      else row.push(d.carryExpired ? 0 : d.prevYearBalance, d.thisYearEntitlement, d.thisYearUtilized, d.currentBalance);
    });
    report.push(row);
  });
  return report;
}
