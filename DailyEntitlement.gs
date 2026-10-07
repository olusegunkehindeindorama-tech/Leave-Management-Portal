/**
 * DAILY ENTITLEMENT + UTILIZATION
 * Requires DailyLeaveModel.gs + loadEntitlementMasterData_ / policies / multipliers.
 */

function dailyShiftMultiplier_(master, empId, dateKey, multFlag) {
  var shifts = master.shifts || master.shiftMap || {};
  var byEmp = shifts[empId] || shifts[String(empId).toUpperCase()] || {};
  var code = String(byEmp[dateKey] || '').trim().toUpperCase();
  if (!code) return 1;
  var flag = String(multFlag || 'No').trim();
  if (flag !== 'Yes' && flag !== 'No') flag = 'No';
  var table = (master.multipliers && master.multipliers[flag]) || {};
  if (table[code] !== undefined && table[code] !== '') return Number(table[code]);
  if (code === 'O') return 0;
  if (code === 'G') return 1;
  if (code === 'A' || code === 'B' || code === 'D' || code === 'N') {
    return flag === 'Yes' ? 1.5 : 1;
  }
  return 1;
}

function recalculateAllLeaveUtilized() {
  return calculateLeaveUtilized();
}

function calculateLeaveUtilized() {
  Logger.log('=== DAILY RECALC START ===');
  var t0 = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) return { success: false, message: 'tblLeave missing' };
  if (!isDailyTblLeaveSchema_(sheet)) {
    return { success: false, message: 'tblLeave is not daily schema. Run migrateTblLeaveToDaily() first.' };
  }

  var master;
  try { master = loadEntitlementMasterData_(); }
  catch (e) { return { success: false, message: 'loadEntitlementMasterData_ failed: ' + e.message }; }
  if (!master.policies || !master.policies.length) {
    return { success: false, message: 'Sys_LeavePolicies empty' };
  }

  if (!master.shifts && !master.shiftMap) master.shifts = loadWideShiftMap_();
  else if (master.shiftMap && !master.shifts) master.shifts = master.shiftMap;
  if (!master.multipliers) {
    try { master.multipliers = loadMultipliersCached_(); }
    catch (e2) { master.multipliers = { Yes: {}, No: {} }; }
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, message: 'No leave rows', updated: 0 };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var iEmp = headers.indexOf('Emp ID');
  var iType = headers.indexOf('Leave Type');
  var iDate = headers.indexOf('Leave Date');
  var iUtil = headers.indexOf('Leave Utilized');
  var iYear = headers.indexOf('Entitlement Year');
  if (iEmp < 0 || iDate < 0 || iUtil < 0 || iYear < 0) {
    return { success: false, message: 'Missing required daily columns' };
  }

  var n = lastRow - 1;
  var data = sheet.getRange(2, 1, n, sheet.getLastColumn()).getValues();
  var items = [];
  for (var r = 0; r < data.length; r++) {
    var emp = String(data[r][iEmp] || '').trim().toUpperCase();
    var dk = dailyDateKey_(data[r][iDate]);
    if (!emp || !dk) continue;
    items.push({
      idx: r, emp: emp,
      type: String(data[r][iType] || '').trim(),
      dateKey: dk,
      dateObj: dailySheetDate_(dk)
    });
  }

  items.sort(function (a, b) {
    if (a.emp !== b.emp) return a.emp < b.emp ? -1 : 1;
    if (a.dateKey !== b.dateKey) return a.dateKey < b.dateKey ? -1 : 1;
    return a.idx - b.idx;
  });

  var remainingPrev = {};
  function getPrev_(empId, calYear) {
    var key = empId + '|' + calYear;
    if (remainingPrev[key] !== undefined) return remainingPrev[key];
    var col = (calYear - 1) + ' Balance';
    var bal = 0;
    if (master.startingBal && master.startingBal[empId]) {
      bal = Number(master.startingBal[empId][col]) || 0;
    }
    remainingPrev[key] = Math.max(0, bal);
    return remainingPrev[key];
  }
  function setPrev_(empId, calYear, v) {
    remainingPrev[empId + '|' + calYear] = Math.max(0, v);
  }

  var utilCol = [];
  var yearCol = [];
  for (var i = 0; i < data.length; i++) {
    utilCol.push([data[i][iUtil]]);
    yearCol.push([data[i][iYear]]);
  }

  var updated = 0;
  for (var j = 0; j < items.length; j++) {
    var it = items[j];
    var empRow = (master.employees && master.employees[it.emp]) || { 'Emp ID': it.emp };
    var pol = matchPolicyForEmp_(master.policies, empRow, it.type, it.dateObj);
    var meta = policyCalcMeta_(pol);
    var multFlag = String((pol && pol['Multiplier']) || 'No').trim();
    var util = dailyShiftMultiplier_(master, it.emp, it.dateKey, multFlag);

    var calYear = Number(it.dateKey.substring(0, 4));
    var year = calYear;
    var usesAnnual = (String(it.type).trim() === 'Annual Leave') ||
      String(meta.deductFrom || '').toLowerCase().indexOf('annual') >= 0;

    if (usesAnnual && meta.deadline) {
      var deadline = carryDeadlineForYear_(meta.deadline, calYear);
      if (deadline) {
        var day = it.dateObj;
        if (day.getTime() <= deadline.getTime()) {
          var rem = getPrev_(it.emp, calYear);
          if (rem > 0 && util > 0) {
            year = calYear - 1;
            setPrev_(it.emp, calYear, rem - util);
          } else if (rem > 0 && util === 0) {
            year = calYear - 1;
          } else {
            year = calYear;
          }
        }
      }
    }

    utilCol[it.idx][0] = util;
    yearCol[it.idx][0] = year;
    updated++;
  }

  sheet.getRange(2, iUtil + 1, n, 1).setValues(utilCol);
  sheet.getRange(2, iYear + 1, n, 1).setValues(yearCol);
  SpreadsheetApp.flush();

  try {
    if (typeof refreshAllEmployeeLeaveColumns_ === 'function') refreshAllEmployeeLeaveColumns_(master);
  } catch (e3) { Logger.log('Balance refresh skipped: ' + e3.message); }
  try { if (typeof cacheClearAll_ === 'function') cacheClearAll_(); } catch (e4) {}

  var ms = new Date().getTime() - t0;
  var msg = 'Daily recalc: updated ' + updated + ' row(s) in ' + ms + ' ms';
  Logger.log('=== DAILY RECALC END === ' + msg);
  return { success: true, message: msg, updated: updated, elapsedMs: ms };
}

function loadWideShiftMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblShift');
  var map = {};
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 2) return map;
  var values = sheet.getDataRange().getValues();
  var display = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var dateKeys = [];
  for (var c = 1; c < values[0].length; c++) {
    dateKeys.push(dailyDateKey_(display[c]) || dailyDateKey_(values[0][c]));
  }
  for (var r = 1; r < values.length; r++) {
    var emp = String(values[r][0] || '').trim().toUpperCase();
    if (!emp) continue;
    if (!map[emp]) map[emp] = {};
    for (var c2 = 1; c2 < values[r].length; c2++) {
      var dk = dateKeys[c2 - 1];
      if (!dk) continue;
      var code = String(values[r][c2] || '').trim().toUpperCase();
      if (code) map[emp][dk] = code;
    }
  }
  Logger.log('loadWideShiftMap_: ' + Object.keys(map).length + ' employees');
  return map;
}
