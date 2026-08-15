/**
 * Fast leave balance for ONE employee (UI path).
 * Applies Deduct from: Casual + Examination usage reduce Annual balance.
 * Uses Cache.gs helpers when available.
 */
function apiGetEmployeeBalance(empId) {
  empId = String(empId || "").trim().toUpperCase();
  if (!empId) return { error: "Employee Not Found" };

  var empMap = (typeof loadEmployeeMapCached_ === "function") ? loadEmployeeMapCached_() : {};
  var emp = empMap[empId];
  if (!emp) return { error: "Employee Not Found" };

  var policies = (typeof loadPoliciesCached_ === "function") ? loadPoliciesCached_() : [];
  var leaveRows = (typeof loadLeaveRowsForEmp_ === "function") ? loadLeaveRowsForEmp_(empId) : [];

  var today = new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;

  var bu = String(emp["Business Unit"] || "").trim();
  var catFull = String(emp["Category"] || "").trim();
  var empStatus = String(emp["Status"] || "").trim();
  var gender = String(emp["Gender"] || "").trim();
  var empName = String(emp["Emp Name"] || "").trim();
  var dept = String(emp["Department"] || "").trim();
  var dojRaw = emp["Date of Join"];
  var doj = (dojRaw instanceof Date) ? dojRaw : parseDDMMYYYY(String(dojRaw || ""));
  var catInitials = mapCategoryToInitials(catFull);

  // Starting bal
  var carryGross = 0;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sb = ss.getSheetByName("StartingBal");
    if (sb) {
      var sbData = sb.getDataRange().getValues();
      var sbH = sbData[0].map(function(h) { return String(h).trim(); });
      var eIdx = sbH.indexOf("Emp No") > -1 ? sbH.indexOf("Emp No") : sbH.indexOf("Emp ID");
      var bIdx = sbH.indexOf(prevYear + " Balance");
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

  // Usage by type/year
  var usage = {}; // type -> { year -> days }
  for (var i = 0; i < leaveRows.length; i++) {
    var lr = leaveRows[i];
    var lt = String(lr["Leave Type"] || "").trim();
    var util = Number(lr["Leave Utilized"]) || 0;
    var ey = Number(lr["Entitlement Year"]) || currentYear;
    if (!lt) continue;
    if (!usage[lt]) usage[lt] = {};
    if (!usage[lt][ey]) usage[lt][ey] = 0;
    usage[lt][ey] += util;
  }

  // Unique leave types from policy that match this employee (entitled)
  var typeMeta = {}; // type -> { entitlement, show, deadline, deductFrom, score }
  for (var p = 0; p < policies.length; p++) {
    var pol = policies[p];
    var lType = String(pol["Leave Type"] || "").trim();
    if (!lType) continue;

    var scoreBU = checkMatch(bu.toUpperCase(), String(pol["Business Unit"] || "").trim().toUpperCase(), 100);
    var scoreCat = checkMatch(catInitials, String(pol["Category"] || "").trim(), 10);
    var scoreStatus = checkMatch(empStatus, String(pol["Status"] || "").trim(), 5);
    var scoreGender = checkMatch(gender, String(pol["Gender"] || "").trim(), 1);
    if (scoreBU === -1 || scoreCat === -1 || scoreStatus === -1 || scoreGender === -1) continue;
    if (!evaluateLifecycle(doj, String(pol["Who is entitled (Lifecycle)"] || "").trim(), today)) continue;

    var totalScore = scoreBU + scoreCat + scoreStatus + scoreGender;
    var entRaw = String(pol["Annual Entitlements"] || "").trim();
    var isUnlimited = (entRaw === "" || entRaw.toLowerCase() === "unlimited");
    var entVal = isUnlimited ? "Unlimited" : (Number(entRaw) || 0);
    var show = String(pol["Balance Page Show"] || "").trim().toLowerCase();
    var showYes = (show === "yes" || show === "y" || show === "true");
    var deductFrom = String(pol["Deduct from"] || "").trim();
    var deadlineRaw = String(pol["Carry Forward Deadline"] || "").trim();
    var deadline = null;
    if (deadlineRaw && deadlineRaw.toLowerCase() !== "no") {
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

  // First pass: raw balances per type
  var annualUsageCurr = 0;
  var annualUsagePrev = 0;
  var casualUsageCurr = 0;
  var examUsageCurr = 0;

  Object.keys(typeMeta).forEach(function(t) {
    var u = usage[t] || {};
    var cy = u[currentYear] || 0;
    var py = u[prevYear] || 0;
    if (t === "Annual Leave") { annualUsageCurr = cy; annualUsagePrev = py; }
    if (t === "Casual Leave") casualUsageCurr = cy;
    if (t === "Examination Leave") examUsageCurr = cy;
  });

  Object.keys(typeMeta).forEach(function(t) {
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
      total = "Unlimited";
      currAvail = "Unlimited";
    } else {
      var baseCarry = (t === "Annual Leave") ? carryGross : 0;
      if (meta.deadline && today > meta.deadline) {
        carryExpired = true;
        prevAvail = 0;
      } else if (t === "Annual Leave") {
        prevAvail = Math.max(0, baseCarry - annualUsagePrev);
      }

      // Deduct-from: Casual & Exam also consume Annual
      var effectiveCurrUsage = cy;
      if (t === "Annual Leave") {
        effectiveCurrUsage = annualUsageCurr + casualUsageCurr + examUsageCurr;
      }

      currAvail = Math.max(0, Number(meta.entitlement) - effectiveCurrUsage);
      // For types that deduct from Annual, balance is min(own remaining, annual remaining)
      if (meta.deductFrom === "Annual Leave") {
        var annualRemaining = 0;
        // Will recompute after annual is known — temporary
        currAvail = Math.max(0, Number(meta.entitlement) - cy);
      }
      total = (typeof currAvail === "number" ? currAvail : 0) + prevAvail;
    }

    balances[t] = total;
    detail[t] = {
      prevYearBalance: meta.isUnlimited ? "Unlimited" : prevAvail,
      thisYearEntitlement: meta.entitlement,
      thisYearUtilized: cy,
      thisYearBalance: currAvail,
      currentBalance: total,
      carryExpired: carryExpired,
      deductFrom: meta.deductFrom,
      show: meta.show
    };
  });

  // Second pass: apply min(casual/exam, annual) for deduct-from types
  var annualBal = balances["Annual Leave"];
  if (typeof annualBal === "number") {
    ["Casual Leave", "Examination Leave"].forEach(function(t) {
      if (!typeMeta[t]) return;
      if (typeMeta[t].deductFrom !== "Annual Leave") return;
      var own = typeof balances[t] === "number" ? balances[t] : 0;
      var capped = Math.min(own, annualBal);
      balances[t] = capped;
      if (detail[t]) {
        detail[t].currentBalance = capped;
        detail[t].thisYearBalance = capped;
      }
    });
  }

  return {
    profile: { id: empId, name: empName, bu: bu, dept: dept, category: catFull },
    balances: balances,
    entitlements: entitlements,
    usage: usageOut,
    detail: detail,
    entitledTypes: Object.keys(typeMeta) // ALL entitled types for dropdown
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
  // Full report: iterate employees (slower; used for CSV only)
  var empMap = loadEmployeeMapCached_();
  var out = {};
  Object.keys(empMap).forEach(function(id) {
    if (id === "_headers") return;
    var emp = empMap[id];
    if (optBU && String(emp["Business Unit"] || "").toUpperCase() !== String(optBU).toUpperCase()) return;
    if (optDept && String(emp["Department"] || "").toUpperCase() !== String(optDept).toUpperCase()) return;
    var bal = apiGetEmployeeBalance(id);
    if (!bal.error) out[id] = bal;
  });
  return out;
}

function generateReportArray(buFilter, deptFilter) {
  var raw = getLeaveBalancesData(null, buFilter, deptFilter);
  var types = {};
  Object.keys(raw).forEach(function(id) {
    Object.keys(raw[id].detail || {}).forEach(function(t) { types[t] = true; });
  });
  var typeList = Object.keys(types).sort();
  var headers = ["Emp ID", "Name", "Department", "BU", "Category"];
  typeList.forEach(function(t) {
    headers.push(t + " — Prev Year Bal");
    headers.push(t + " — This Year Entitlement");
    headers.push(t + " — Utilized This Year");
    headers.push(t + " — Current Balance");
  });
  var report = [headers];
  Object.keys(raw).forEach(function(id) {
    var emp = raw[id];
    var row = [emp.profile.id, emp.profile.name, emp.profile.dept, emp.profile.bu, emp.profile.category];
    typeList.forEach(function(t) {
      var d = (emp.detail && emp.detail[t]) ? emp.detail[t] : null;
      if (!d) row.push("", "", "", "");
      else row.push(d.carryExpired ? 0 : d.prevYearBalance, d.thisYearEntitlement, d.thisYearUtilized, d.currentBalance);
    });
    report.push(row);
  });
  return report;
}
