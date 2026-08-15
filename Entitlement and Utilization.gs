/**
 * LEAVE ENTITLEMENT & UTILIZATION
 * Official table (leave calculation.docx):
 *   ShiftRoaster: G=1, A/B/D/N=1.5, O=0
 *   ActualDays: every calendar day = 1 (including O)
 * No roster → Mon–Fri = G, Sat–Sun = O
 */

function checkMatch(empValue, policyValue, weight) {
  if (policyValue === "DEFAULT") return 0;
  if (String(empValue).toLowerCase() === String(policyValue).toLowerCase()) return weight;
  return -1;
}

function mapCategoryToInitials(fullCategory) {
  var cat = String(fullCategory || "").toLowerCase();
  if (cat.indexOf("junior staff") > -1) return "JS";
  if (cat.indexOf("non union mgt") > -1) return "NUMS";
  if (cat.indexOf("non union senior") > -1) return "NUSS";
  if (cat.indexOf("senior staff") > -1) return "SS";
  if (cat.indexOf("trainee") > -1) return "T";
  return fullCategory;
}

function parseDDMMYYYY(dateString) {
  if (!dateString) return new Date();
  if (dateString instanceof Date) return dateString;
  var parts = String(dateString).split("/");
  if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
  return new Date(dateString);
}

function evaluateLifecycle(doj, ruleString, todayDate) {
  if (!ruleString || ruleString === "DEFAULT" || ruleString === "") return true;
  var regex = /([><=]+)\s*(\d+)\s*(year|years|month|months|day|days)/i;
  var match = String(ruleString).match(regex);
  if (!match) return true;
  var operator = match[1];
  var value = parseFloat(match[2]);
  var unit = match[3].toLowerCase();
  var diffDays = (todayDate.getTime() - doj.getTime()) / (1000 * 3600 * 24);
  var targetDays = 0;
  if (unit.indexOf("year") > -1) targetDays = value * 365.25;
  else if (unit.indexOf("month") > -1) targetDays = value * 30.4375;
  else targetDays = value;
  switch (operator) {
    case ">=": return diffDays >= targetDays;
    case ">": return diffDays > targetDays;
    case "<=": return diffDays <= targetDays;
    case "<": return diffDays < targetDays;
    case "==":
    case "=": return Math.round(diffDays) === Math.round(targetDays);
    default: return true;
  }
}

function formatDateKey(dateObj) {
  var y = dateObj.getFullYear();
  var m = ("0" + (dateObj.getMonth() + 1)).slice(-2);
  var d = ("0" + dateObj.getDate()).slice(-2);
  return y + "-" + m + "-" + d;
}

/** Calculation Method from Sys_LeavePolicies for this leave type. */
function getLeaveCalcMeta_(leaveType) {
  var policies = (typeof loadPoliciesCached_ === "function") ? loadPoliciesCached_() : [];
  var method = "ActualDays";
  var deductFrom = "";
  for (var i = 0; i < policies.length; i++) {
    if (String(policies[i]["Leave Type"] || "").trim() !== String(leaveType).trim()) continue;
    var m = String(policies[i]["Calculation Method"] || "").trim();
    if (m) method = m;
    deductFrom = String(policies[i]["Deduct from"] || "").trim();
    break;
  }
  var ml = method.toLowerCase().replace(/\s+/g, "");
  if (ml.indexOf("shift") > -1 || ml.indexOf("roaster") > -1 || ml.indexOf("roster") > -1) {
    method = "ShiftRoaster";
  } else {
    method = "ActualDays";
  }
  return { method: method, deductFrom: deductFrom };
}

/** Official Shift Roaster multipliers: G=1, A/B/D/N(/M)=1.5, O=0 */
function shiftRoasterWeight_(code) {
  code = String(code || "").trim().toUpperCase();
  if (code === "O" || code === "OFF" || code === "REST" || code === "R" ||
      code === "WO" || code === "W/O" || code === "LEAVE" || code === "L") return 0;
  if (code === "G") return 1;
  if (["A", "B", "D", "N", "M"].indexOf(code) > -1) return 1.5;
  return 1;
}

/** No roster: Mon–Fri = G, Sat–Sun = O */
function defaultShiftCode_(dateObj) {
  var day = dateObj.getDay();
  return (day === 0 || day === 6) ? "O" : "G";
}

/**
 * Core utilization for a date range.
 * Example: D,D,O,O,G → ShiftRoaster 4 / ActualDays 5
 */
function calculateLeaveUtilize(empId, startDate, endDate, leaveType) {
  var sDate = new Date(startDate);
  var eDate = new Date(endDate);
  if (isNaN(sDate.getTime()) || isNaN(eDate.getTime()) || eDate < sDate) return 0;

  var meta = getLeaveCalcMeta_(leaveType || "Annual Leave");
  var inclusive = Math.round((eDate.getTime() - sDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;

  if (meta.method === "ActualDays") return inclusive;

  var shiftMap = (typeof loadShiftMapForEmp_ === "function") ? loadShiftMapForEmp_(empId) : {};
  var total = 0;
  var curr = new Date(sDate.getFullYear(), sDate.getMonth(), sDate.getDate());
  var end = new Date(eDate.getFullYear(), eDate.getMonth(), eDate.getDate());

  while (curr <= end) {
    var key = formatDateKey(curr);
    var code = shiftMap[key];
    if (!code) code = defaultShiftCode_(curr);
    total += shiftRoasterWeight_(code);
    curr.setDate(curr.getDate() + 1);
  }
  return total;
}

/**
 * Recalculate Leave Utilized (and No of Days) for every tblLeave row
 * using the corrected ShiftRoaster / ActualDays logic.
 * Safe for web app — returns JSON (no SpreadsheetApp.getUi).
 */
function calculateLeaveUtilized() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leaveSheet = ss.getSheetByName("tblLeave");
  if (!leaveSheet) {
    return { success: false, message: "tblLeave sheet missing.", updated: 0 };
  }

  var data = leaveSheet.getDataRange().getValues();
  if (data.length < 2) {
    return { success: true, message: "No leave records to recalculate.", updated: 0 };
  }

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var empIdx = headers.indexOf("Emp ID");
  var typeIdx = headers.indexOf("Leave Type");
  var startIdx = headers.indexOf("Start Date");
  var endIdx = headers.indexOf("End Date");
  var utilIdx = headers.indexOf("Leave Utilized");
  var daysIdx = headers.indexOf("No of Days");
  var yearIdx = headers.indexOf("Entitlement Year");

  if (empIdx < 0 || startIdx < 0 || endIdx < 0 || typeIdx < 0) {
    return { success: false, message: "tblLeave is missing required columns.", updated: 0 };
  }

  var updated = 0;
  var skipped = 0;

  // Batch writes for speed: collect values then set in columns
  var utilCol = [];
  var daysCol = [];
  var yearCol = [];
  var yearUpdates = []; // {row, val} only when empty

  for (var i = 1; i < data.length; i++) {
    var empId = String(data[i][empIdx]).trim();
    var s = new Date(data[i][startIdx]);
    var e = new Date(data[i][endIdx]);

    if (!empId || isNaN(s.getTime()) || isNaN(e.getTime())) {
      skipped++;
      if (utilIdx > -1) utilCol.push([data[i][utilIdx]]);
      if (daysIdx > -1) daysCol.push([data[i][daysIdx]]);
      continue;
    }

    var lt = String(data[i][typeIdx]).trim();
    var util = calculateLeaveUtilize(empId, s, e, lt);
    var calDays = Math.round((e.getTime() - s.getTime()) / (1000 * 60 * 60 * 24)) + 1;

    if (utilIdx > -1) utilCol.push([util]);
    if (daysIdx > -1) daysCol.push([calDays]);
    if (yearIdx > -1 && (data[i][yearIdx] === "" || data[i][yearIdx] == null)) {
      yearUpdates.push({ row: i + 1, val: s.getFullYear() });
    }
    updated++;
  }

  if (utilIdx > -1 && utilCol.length) {
    leaveSheet.getRange(2, utilIdx + 1, utilCol.length, 1).setValues(utilCol);
  }
  if (daysIdx > -1 && daysCol.length) {
    leaveSheet.getRange(2, daysIdx + 1, daysCol.length, 1).setValues(daysCol);
  }
  for (var y = 0; y < yearUpdates.length; y++) {
    leaveSheet.getRange(yearUpdates[y].row, yearIdx + 1).setValue(yearUpdates[y].val);
  }

  if (typeof cacheClearAll_ === "function") cacheClearAll_();
  // Also drop per-emp leave caches so balances refresh
  try {
    CacheService.getScriptCache().removeAll(
      Object.keys(CacheService.getScriptCache().getAll([]) || {})
    );
  } catch (e) {}

  return {
    success: true,
    message: "Recalculated Leave Utilized for " + updated + " record(s)" +
      (skipped ? " (" + skipped + " skipped — missing ID/dates)." : ".") +
      " Uses Shift Roaster multipliers or Actual Days per leave type policy.",
    updated: updated,
    skipped: skipped
  };
}

/** Alias for UI / scripts that call the longer name */
function recalculateAllLeaveUtilized() {
  return calculateLeaveUtilized();
}

function generateEntitlementMatrix() {
  return { success: false, message: "Use balance report from the web app for live entitlements." };
}
