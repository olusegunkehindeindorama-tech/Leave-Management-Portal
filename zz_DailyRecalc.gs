/**
 * zz_DailyRecalc.gs — loads last; daily util + entitlement year.
 * Carry-forward: deadline day is INCLUSIVE (last allowable prev-year day).
 * Leave Date is noon; deadline is midnight — compare calendar days only.
 */

function calculateLeaveUtilizedDaily_() {
  Logger.log('=== DAILY RECALC START ===');
  var t0 = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) return { success: false, message: 'tblLeave missing' };
  if (typeof isDailyTblLeaveSchema_ === 'function' && !isDailyTblLeaveSchema_(sheet)) {
    return { success: false, message: 'tblLeave is not daily schema. Run migrateTblLeaveToDaily() first.' };
  }

  var master;
  try { master = loadEntitlementMasterData_(); }
  catch (e) { return { success: false, message: 'loadEntitlementMasterData_ failed: ' + e.message }; }
  if (!master.policies || !master.policies.length) {
    return { success: false, message: 'Sys_LeavePolicies empty' };
  }
  if (!master.multipliers) master.multipliers = { Yes: {}, No: {} };
  if (!master.shifts || !Object.keys(master.shifts).length) {
    if (typeof loadWideShiftMap_ === 'function') master.shifts = loadWideShiftMap_();
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
    return { success: false, message: 'Missing required daily columns (need Leave Date)' };
  }

  var n = lastRow - 1;
  var data = sheet.getRange(2, 1, n, sheet.getLastColumn()).getValues();
  var items = [];
  for (var r = 0; r < data.length; r++) {
    var emp = String(data[r][iEmp] || '').trim().toUpperCase();
    var dk = (typeof dailyDateKey_ === 'function')
      ? dailyDateKey_(data[r][iDate])
      : formatDateKey(new Date(data[r][iDate]));
    if (!emp || !dk) continue;
    var dObj = (typeof dailySheetDate_ === 'function')
      ? dailySheetDate_(dk)
      : new Date(Number(dk.substring(0, 4)), Number(dk.substring(5, 7)) - 1, Number(dk.substring(8, 10)), 12, 0, 0);
    items.push({ idx: r, emp: emp, type: String(data[r][iType] || '').trim(), dateKey: dk, dateObj: dObj });
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

  function utilForDay_(empId, dateKey, dateObj, leaveType) {
    var empRow = (master.employees && master.employees[empId]) || { 'Emp ID': empId };
    var pol = matchPolicyForEmp_(master.policies, empRow, leaveType, dateObj);
    var meta = policyCalcMeta_(pol);

    if (meta.method === 'ActualDays') {
      return { util: 1, meta: meta, shift: '' };
    }

    var shiftMap = (master.shifts && master.shifts[empId]) || {};
    var code = String(shiftMap[dateKey] || '').trim().toUpperCase();
    if (!code) {
      code = (typeof defaultShiftCode_ === 'function')
        ? defaultShiftCode_(dateObj)
        : ((dateObj.getDay() === 0 || dateObj.getDay() === 6) ? 'O' : 'G');
    }

    var w;
    if (typeof multiplierWeight_ === 'function') {
      w = multiplierWeight_(master.multipliers, meta.multFlag, code);
    } else {
      var table = (meta.multFlag === 'Yes') ? (master.multipliers.Yes || {}) : (master.multipliers.No || {});
      w = (table[code] !== undefined && !isNaN(Number(table[code]))) ? Number(table[code]) : (code === 'O' ? 0 : 1);
    }
    if (isNaN(w) || w < 0) w = 0;
    w = Math.round(w * 1000) / 1000;
    return { util: w, meta: meta, shift: code };
  }

  var utilCol = [];
  var yearCol = [];
  for (var i = 0; i < data.length; i++) {
    utilCol.push([0]);
    yearCol.push(['']);
  }

  var updated = 0;
  var sampleLog = 0;
  for (var j = 0; j < items.length; j++) {
    var it = items[j];
    var res = utilForDay_(it.emp, it.dateKey, it.dateObj, it.type);
    var util = res.util;
    var meta = res.meta;
    var calYear = Number(it.dateKey.substring(0, 4));
    var year = calYear;

    var usesAnnual = (String(it.type).trim() === 'Annual Leave') ||
      String(meta.deductFrom || '').toLowerCase().indexOf('annual') >= 0;

    if (usesAnnual && meta.deadline) {
      var deadline = carryDeadlineForYear_(meta.deadline, calYear);
      // Calendar-day compare: cutoff day is INCLUSIVE (last day for prev-year balance).
      // Leave Date is stored at noon; deadline is midnight — do not compare raw getTime().
      if (deadline) {
        var leaveDay = new Date(it.dateObj.getFullYear(), it.dateObj.getMonth(), it.dateObj.getDate());
        var deadlineDay = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
        if (leaveDay.getTime() <= deadlineDay.getTime()) {
          var rem = getPrev_(it.emp, calYear);
          if (rem > 0) {
            year = calYear - 1;
            if (util > 0) setPrev_(it.emp, calYear, rem - util);
          } else {
            year = calYear;
          }
        }
      }
    }

    utilCol[it.idx][0] = util;
    yearCol[it.idx][0] = year;
    updated++;

    if (sampleLog < 8) {
      Logger.log('sample ' + it.emp + ' ' + it.dateKey + ' type=' + it.type +
        ' shift=' + (res.shift || '') + ' util=' + util + ' year=' + year +
        ' method=' + meta.method + ' multFlag=' + meta.multFlag);
      sampleLog++;
    }
  }

  sheet.getRange(2, iUtil + 1, n, 1).setValues(utilCol);
  sheet.getRange(2, iYear + 1, n, 1).setValues(yearCol);
  SpreadsheetApp.flush();

  try {
    if (typeof refreshAllEmployeeLeaveColumns_ === 'function') refreshAllEmployeeLeaveColumns_(master);
  } catch (e3) { Logger.log('Balance refresh skipped: ' + e3.message); }
  try { if (typeof cacheClearAll_ === 'function') cacheClearAll_(); } catch (e4) {}

  var ms = new Date().getTime() - t0;
  var msg = 'Daily recalc: updated ' + updated + ' row(s) in ' + ms + ' ms (util from policy/shift only)';
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
    var dk = (typeof dailyDateKey_ === 'function')
      ? (dailyDateKey_(display[c]) || dailyDateKey_(values[0][c]))
      : String(display[c] || values[0][c] || '').substring(0, 10);
    dateKeys.push(dk);
  }
  for (var r = 1; r < values.length; r++) {
    var emp = String(values[r][0] || '').trim().toUpperCase();
    if (!emp) continue;
    if (!map[emp]) map[emp] = {};
    for (var c2 = 1; c2 < values[r].length; c2++) {
      var key = dateKeys[c2 - 1];
      if (!key) continue;
      var code = String(values[r][c2] || '').trim().toUpperCase();
      if (code) map[emp][key] = code;
    }
  }
  Logger.log('loadWideShiftMap_: ' + Object.keys(map).length + ' employees');
  return map;
}

function calculateLeaveUtilized() {
  Logger.log('zz_DailyRecalc: calculateLeaveUtilized → daily');
  return calculateLeaveUtilizedDaily_();
}
function recalculateAllLeaveUtilized() {
  Logger.log('zz_DailyRecalc: recalculateAllLeaveUtilized → daily');
  return calculateLeaveUtilizedDaily_();
}
