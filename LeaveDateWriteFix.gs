/**
 * LeaveDateWriteFix.gs — load AFTER Leave Cleanup.gs, DateUtils.gs, CarryForwardFix.gs
 *
 * 1. After cleanup pipeline / recalculate, force every Start/End on tblLeave
 *    to local noon + dd-mmm-yyyy (stops WAT 23:00 previous-day drift).
 * 2. Step-by-step Logger logs for pipeline and normalize.
 * 3. Overrides leaveParseDate_ to use toCalendarDate_ (2h nudge) when available.
 */

leaveParseDate_ = function (val) {
  if (typeof toCalendarDate_ === 'function') {
    var cd = toCalendarDate_(val);
    if (cd) return cd;
  }
  if (val instanceof Date && !isNaN(val.getTime())) {
    var d0 = new Date(val.getTime());
    if (d0.getHours() >= 20) {
      d0 = new Date(d0.getTime() + 2 * 60 * 60 * 1000);
    }
    return new Date(d0.getFullYear(), d0.getMonth(), d0.getDate());
  }
  if (val === null || val === undefined || val === '') return null;
  var s = String(val).trim();
  var months = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
  if (mon && months[mon[2].toLowerCase()] !== undefined) {
    var y = Number(mon[3]); if (y < 100) y = y >= 70 ? 1900 + y : 2000 + y;
    return new Date(y, months[mon[2].toLowerCase()], Number(mon[1]));
  }
  var slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (slash) {
    var yy = Number(slash[3]); if (yy < 100) yy = yy >= 70 ? 1900 + yy : 2000 + yy;
    return new Date(yy, Number(slash[2]) - 1, Number(slash[1]));
  }
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  var dt = new Date(s);
  if (!isNaN(dt.getTime())) {
    if (dt.getHours() >= 20) dt = new Date(dt.getTime() + 2 * 60 * 60 * 1000);
    return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  }
  return null;
};

function leaveSheetDate_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof toSheetDateValue_ === 'function') {
    var v = toSheetDateValue_(val);
    if (v !== '' && v !== null && v !== undefined) return v;
  }
  var d = leaveParseDate_(val);
  if (!d) return '';
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
}

function leaveWriteDateCell_(sheet, row, col, val) {
  var v = leaveSheetDate_(val);
  if (v === '' || v === null) return;
  sheet.getRange(row, col).setValue(v).setNumberFormat('dd-mmm-yyyy');
}

/**
 * Walk tblLeave and rewrite every Start/End as noon of the intended calendar day.
 */
function normalizeAllTblLeaveStartEndDates() {
  Logger.log('=== normalizeAllTblLeaveStartEndDates START ===');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) {
    Logger.log('normalize: tblLeave missing');
    return { success: false, message: 'tblLeave missing' };
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    Logger.log('normalize: no data rows');
    return { success: true, message: 'no rows', fixed: 0 };
  }
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  if (startIdx < 0 || endIdx < 0) {
    Logger.log('normalize: Start/End columns missing');
    return { success: false, message: 'Start/End columns missing' };
  }

  var fixed = 0;
  for (var i = 1; i < data.length; i++) {
    var rawS = data[i][startIdx];
    var rawE = data[i][endIdx];
    if (rawS === '' && rawE === '') continue;

    var newS = leaveSheetDate_(rawS);
    var newE = leaveSheetDate_(rawE);
    var row = i + 1;
    var changed = false;

    if (newS) {
      var needS = true;
      if (rawS instanceof Date && !isNaN(rawS.getTime())) {
        if (rawS.getFullYear() === newS.getFullYear() &&
            rawS.getMonth() === newS.getMonth() &&
            rawS.getDate() === newS.getDate() &&
            rawS.getHours() === 12 && rawS.getMinutes() === 0) {
          needS = false;
          sheet.getRange(row, startIdx + 1).setNumberFormat('dd-mmm-yyyy');
        }
      }
      if (needS) {
        leaveWriteDateCell_(sheet, row, startIdx + 1, newS);
        changed = true;
      }
    }
    if (newE) {
      var needE = true;
      if (rawE instanceof Date && !isNaN(rawE.getTime())) {
        if (rawE.getFullYear() === newE.getFullYear() &&
            rawE.getMonth() === newE.getMonth() &&
            rawE.getDate() === newE.getDate() &&
            rawE.getHours() === 12 && rawE.getMinutes() === 0) {
          needE = false;
          sheet.getRange(row, endIdx + 1).setNumberFormat('dd-mmm-yyyy');
        }
      }
      if (needE) {
        leaveWriteDateCell_(sheet, row, endIdx + 1, newE);
        changed = true;
      }
    }
    if (changed) fixed++;
    if (i % 1000 === 0) {
      Logger.log('normalize progress: row ' + row + ' fixed so far ' + fixed);
      SpreadsheetApp.flush();
    }
  }
  SpreadsheetApp.flush();
  var msg = 'normalizeAllTblLeaveStartEndDates: fixed ' + fixed + ' row(s) of ' + (lastRow - 1);
  Logger.log('=== normalizeAllTblLeaveStartEndDates END === ' + msg);
  return { success: true, message: msg, fixed: fixed, total: lastRow - 1 };
}

(function () {
  if (typeof runLeaveCleanupPipeline === 'function') {
    var _origPipeline = runLeaveCleanupPipeline;
    runLeaveCleanupPipeline = function () {
      Logger.log('=== PIPELINE START (wrapped) ===');
      Logger.log('[1] original pipeline (overlap + dedupe + recalc) ...');
      var result = _origPipeline();
      Logger.log('[1] pipeline core done: ' + (result && result.message ? result.message : ''));
      Logger.log('[2] normalizeAllTblLeaveStartEndDates (force noon) ...');
      var norm = normalizeAllTblLeaveStartEndDates();
      Logger.log('[2] normalize done: ' + (norm && norm.message ? norm.message : ''));
      Logger.log('=== PIPELINE END (wrapped) ===');
      if (result) {
        result.normalize = norm;
        result.message = (result.message || '') + ' | ' + (norm.message || '');
      }
      return result;
    };
    Logger.log('LeaveDateWriteFix: runLeaveCleanupPipeline wrapped with noon normalize');
  }

  if (typeof calculateLeaveUtilized === 'function') {
    var _origRecalc = calculateLeaveUtilized;
    calculateLeaveUtilized = function () {
      Logger.log('=== RECALC START (wrapped) ===');
      var result = _origRecalc();
      Logger.log('=== RECALC core done === ' + (result && result.message ? result.message : ''));
      Logger.log('=== RECALC normalize dates ===');
      var norm = normalizeAllTblLeaveStartEndDates();
      Logger.log('=== RECALC END === ' + (norm && norm.message ? norm.message : ''));
      if (result) {
        result.normalize = norm;
        result.message = (result.message || '') + ' | ' + (norm.message || '');
      }
      return result;
    };
    recalculateAllLeaveUtilized = function () {
      return calculateLeaveUtilized();
    };
    Logger.log('LeaveDateWriteFix: calculateLeaveUtilized wrapped with noon normalize');
  }
})();
