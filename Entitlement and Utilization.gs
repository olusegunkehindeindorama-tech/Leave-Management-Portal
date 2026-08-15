/**
 * LEAVE ENTITLEMENT & UTILIZATION
 * - ShiftRoaster + Multiplier (Yes → A/B/D/N/M = 1.5)
 * - No roster day → Mon–Fri = G (1), Sat–Sun = O (0)
 * - ActualDays = inclusive calendar days
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

/**
 * Resolve calculation method + multiplier flag for a leave type from policies.
 */
function getLeaveCalcMeta_(leaveType) {
  var policies = (typeof loadPoliciesCached_ === "function") ? loadPoliciesCached_() : [];
  var method = "ActualDays";
  var useMultiplier = false;
  var deductFrom = "";
  for (var i = 0; i < policies.length; i++) {
    if (String(policies[i]["Leave Type"] || "").trim() !== String(leaveType).trim()) continue;
    var m = String(policies[i]["Calculation Method"] || "").trim();
    if (m) method = m;
    var mult = String(policies[i]["Multiplier"] || "").trim().toLowerCase();
    useMultiplier = (mult === "yes" || mult === "y" || mult === "true");
    deductFrom = String(policies[i]["Deduct from"] || "").trim();
    break;
  }
  return { method: method, useMultiplier: useMultiplier, deductFrom: deductFrom };
}

/**
 * Day weight for one shift code.
 */
function shiftDayWeight_(code, useMultiplier) {
  code = String(code || "").trim().toUpperCase();
  if (code === "O" || code === "OFF" || code === "REST" || code === "R" || code === "WO" || code === "W/O" || code === "LEAVE" || code === "L") return 0;
  if (code === "G") return 1;
  if (["A", "B", "D", "N", "M"].indexOf(code) > -1) return useMultiplier ? 1.5 : 1;
  // Unknown code: treat as working day
  return 1;
}

/**
 * Default code when no roster: Mon–Fri = G, Sat–Sun = O
 */
function defaultShiftCode_(dateObj) {
  var day = dateObj.getDay(); // 0=Sun ... 6=Sat
  return (day === 0 || day === 6) ? "O" : "G";
}

/**
 * Core utilization for a date range.
 * @param {string} empId
 * @param {Date|string} startDate
 * @param {Date|string} endDate
 * @param {string} [leaveType] - used to pick Calculation Method + Multiplier
 */
function calculateLeaveUtilize(empId, startDate, endDate, leaveType) {
  var sDate = new Date(startDate);
  var eDate = new Date(endDate);
  if (isNaN(sDate.getTime()) || isNaN(eDate.getTime()) || eDate < sDate) return 0;

  var meta = getLeaveCalcMeta_(leaveType || "Annual Leave");
  var inclusive = Math.round((eDate.getTime() - sDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;

  if (meta.method === "ActualDays") return inclusive;

  // ShiftRoaster path
  var shiftMap = (typeof loadShiftMapForEmp_ === "function") ? loadShiftMapForEmp_(empId) : {};
  var total = 0;
  var curr = new Date(sDate.getFullYear(), sDate.getMonth(), sDate.getDate());
  var end = new Date(eDate.getFullYear(), eDate.getMonth(), eDate.getDate());

  while (curr <= end) {
    var key = formatDateKey(curr);
    var code = shiftMap[key];
    if (!code) code = defaultShiftCode_(curr);
    total += shiftDayWeight_(code, meta.useMultiplier);
    curr.setDate(curr.getDate() + 1);
  }
  return total;
}

/**
 * Batch recalculation for all tblLeave (admin / scheduled).
 * Preserves prior split/carry logic where possible.
 */
function calculateLeaveUtilized() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leaveSheet = ss.getSheetByName("tblLeave");
  if (!leaveSheet) {
    SpreadsheetApp.getUi().alert("tblLeave missing");
    return;
  }
  var data = leaveSheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var empIdx = headers.indexOf("Emp ID");
  var typeIdx = headers.indexOf("Leave Type");
  var startIdx = headers.indexOf("Start Date");
  var endIdx = headers.indexOf("End Date");
  var utilIdx = headers.indexOf("Leave Utilized");
  var daysIdx = headers.indexOf("No of Days");
  var yearIdx = headers.indexOf("Entitlement Year");

  for (var i = 1; i < data.length; i++) {
    var empId = String(data[i][empIdx]).trim();
    if (!empId) continue;
    var s = new Date(data[i][startIdx]);
    var e = new Date(data[i][endIdx]);
    if (isNaN(s.getTime()) || isNaN(e.getTime())) continue;
    var lt = String(data[i][typeIdx]).trim();
    var util = calculateLeaveUtilize(empId, s, e, lt);
    var calDays = Math.round((e.getTime() - s.getTime()) / (1000 * 60 * 60 * 24)) + 1;
    if (utilIdx > -1) leaveSheet.getRange(i + 1, utilIdx + 1).setValue(util);
    if (daysIdx > -1) leaveSheet.getRange(i + 1, daysIdx + 1).setValue(calDays);
    if (yearIdx > -1 && !data[i][yearIdx]) leaveSheet.getRange(i + 1, yearIdx + 1).setValue(s.getFullYear());
  }
  if (typeof cacheClearAll_ === "function") cacheClearAll_();
  SpreadsheetApp.getUi().alert("Leave Utilized recalculated for all rows.");
}

function generateEntitlementMatrix() {
  SpreadsheetApp.getUi().alert("Use balance report from the web app for live entitlements.");
}
