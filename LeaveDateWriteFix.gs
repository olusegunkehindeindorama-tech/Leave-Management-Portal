/**
 * LeaveDateWriteFix.gs — load LAST (after Leave Cleanup, DateUtils, CarryForwardFix)
 *
 * Goals:
 *  1. Start/End stay calendar-correct (noon local, format dd-mmm-yyyy).
 *  2. Pipeline runs normalize ONCE (batch writes) — not after every step.
 *  3. Dedupe does NOT rewrite every date cell (that was ~14k setValue calls).
 *  4. Recalc does NOT normalize when already inside the pipeline.
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

function leaveIsNoonCalendar_(raw, noonDate) {
  if (!(raw instanceof Date) || isNaN(raw.getTime()) || !noonDate) return false;
  return raw.getFullYear() === noonDate.getFullYear() &&
    raw.getMonth() === noonDate.getMonth() &&
    raw.getDate() === noonDate.getDate() &&
    raw.getHours() === 12 &&
    raw.getMinutes() === 0;
}

/**
 * Batch-normalize all Start/End on tblLeave.
 * Uses setValues on whole columns — NOT per-cell setValue.
 */
function normalizeAllTblLeaveStartEndDates() {
  var t0 = new Date().getTime();
  Logger.log('=== normalizeAllTblLeaveStartEndDates START (batch) ===');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) {
    Logger.log('normalize: tblLeave missing');
    return { success: false, message: 'tblLeave missing' };
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    return { success: true, message: 'no rows', fixed: 0 };
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  if (startIdx < 0 || endIdx < 0) {
    return { success: false, message: 'Start/End columns missing' };
  }

  var startCol = sheet.getRange(2, startIdx + 1, lastRow - 1, 1).getValues();
  var endCol = sheet.getRange(2, endIdx + 1, lastRow - 1, 1).getValues();
  var n = startCol.length;
  var fixed = 0;
  var outS = [];
  var outE = [];
  var needWrite = false;

  for (var i = 0; i < n; i++) {
    var rawS = startCol[i][0];
    var rawE = endCol[i][0];
    var newS = (rawS === '' || rawS === null) ? rawS : leaveSheetDate_(rawS);
    var newE = (rawE === '' || rawE === null) ? rawE : leaveSheetDate_(rawE);

    var changeS = false;
    var changeE = false;
    if (newS && !leaveIsNoonCalendar_(rawS, newS)) {
      changeS = true;
      fixed++;
    } else if (newS) {
      newS = rawS;
    }
    if (newE && !leaveIsNoonCalendar_(rawE, newE)) {
      changeE = true;
      if (!changeS) fixed++;
    } else if (newE) {
      newE = rawE;
    }
    if (changeS || changeE) needWrite = true;
    outS.push([newS === '' || newS === null ? rawS : newS]);
    outE.push([newE === '' || newE === null ? rawE : newE]);
  }

  if (needWrite) {
    sheet.getRange(2, startIdx + 1, n, 1).setValues(outS);
    sheet.getRange(2, endIdx + 1, n, 1).setValues(outE);
    sheet.getRange(2, startIdx + 1, n, 1).setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(2, endIdx + 1, n, 1).setNumberFormat('dd-mmm-yyyy');
    SpreadsheetApp.flush();
  }

  var ms = new Date().getTime() - t0;
  var msg = 'normalize batch: fixed ~' + fixed + ' of ' + n + ' rows in ' + ms + ' ms' +
    (needWrite ? '' : ' (already noon — no write)');
  Logger.log('=== normalizeAllTblLeaveStartEndDates END === ' + msg);
  return { success: true, message: msg, fixed: fixed, total: n, elapsedMs: ms };
}

(function () {
  if (typeof exactDedupeLeaveRecordsInPlace_ === 'function') {
    var _origDedupe = exactDedupeLeaveRecordsInPlace_;
    exactDedupeLeaveRecordsInPlace_ = function () {
      Logger.log('exactDedupe (dates NOT rewritten — normalize runs once later)');
      var savedWrite = (typeof leaveWriteDateCell_ === 'function') ? leaveWriteDateCell_ : null;
      leaveWriteDateCell_ = function () { /* no-op during dedupe */ };
      var result;
      try {
        result = _origDedupe();
      } finally {
        if (savedWrite) leaveWriteDateCell_ = savedWrite;
      }
      return result;
    };
    Logger.log('LeaveDateWriteFix: exactDedupe patched — skips date cell writes');
  }
})();

var _LEAVE_IN_PIPELINE_ = false;

(function () {
  if (typeof runLeaveCleanupPipeline === 'function') {
    var _origPipeline = runLeaveCleanupPipeline;
    runLeaveCleanupPipeline = function () {
      Logger.log('=== PIPELINE START (optimized) ===');
      _LEAVE_IN_PIPELINE_ = true;
      var result;
      try {
        Logger.log('[1] overlap + dedupe + recalc (no mid-normalize) ...');
        result = _origPipeline();
        Logger.log('[1] core done: ' + (result && result.message ? result.message : ''));
      } finally {
        _LEAVE_IN_PIPELINE_ = false;
      }
      Logger.log('[2] single batch normalizeAllTblLeaveStartEndDates ...');
      var norm = normalizeAllTblLeaveStartEndDates();
      Logger.log('[2] done: ' + (norm && norm.message ? norm.message : ''));
      Logger.log('=== PIPELINE END (optimized) ===');
      if (result) {
        result.normalize = norm;
        result.message = (result.message || '') + ' | ' + (norm.message || '');
      }
      return result;
    };
    Logger.log('LeaveDateWriteFix: pipeline wraps once with batch normalize at end');
  }

  if (typeof calculateLeaveUtilized === 'function') {
    var _origRecalc = calculateLeaveUtilized;
    calculateLeaveUtilized = function () {
      Logger.log('=== RECALC START === (inPipeline=' + _LEAVE_IN_PIPELINE_ + ')');
      var result = _origRecalc();
      Logger.log('=== RECALC core done === ' + (result && result.message ? result.message : ''));
      if (!_LEAVE_IN_PIPELINE_) {
        Logger.log('=== RECALC standalone → batch normalize ===');
        var norm = normalizeAllTblLeaveStartEndDates();
        if (result) {
          result.normalize = norm;
          result.message = (result.message || '') + ' | ' + (norm.message || '');
        }
        Logger.log('=== RECALC END === ' + (norm && norm.message ? norm.message : ''));
      } else {
        Logger.log('=== RECALC END === skip normalize (pipeline will do it once)');
      }
      return result;
    };
    recalculateAllLeaveUtilized = function () {
      return calculateLeaveUtilized();
    };
    Logger.log('LeaveDateWriteFix: recalculate normalizes only when standalone');
  }
})();
