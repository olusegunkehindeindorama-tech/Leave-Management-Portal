/** Daily leave cleanup — dedupe by Emp|LeaveDate|LeaveType only. */

function runLeaveCleanupPipeline() {
  Logger.log('=== DAILY PIPELINE START ===');
  var t0 = new Date().getTime();
  var dedupe = exactDedupeDailyLeaveInPlace_();
  Logger.log('[1] dedupe: ' + (dedupe.message || ''));
  var recalc = null;
  if (typeof calculateLeaveUtilized === 'function') {
    recalc = calculateLeaveUtilized();
    Logger.log('[2] recalc: ' + (recalc && recalc.message ? recalc.message : ''));
  }
  var ms = new Date().getTime() - t0;
  var msg = 'Daily pipeline ' + ms + ' ms | ' + (dedupe.message || '') +
    (recalc && recalc.message ? ' | ' + recalc.message : '');
  Logger.log('=== DAILY PIPELINE END === ' + msg);
  return { success: true, message: msg, dedupe: dedupe, recalc: recalc, elapsedMs: ms };
}

function exactDedupeDailyLeaveInPlace_() {
  Logger.log('exactDedupeDailyLeaveInPlace_: begin');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) return { success: false, message: 'tblLeave missing' };
  if (!isDailyTblLeaveSchema_(sheet)) {
    return { success: false, message: 'Not daily schema — migrate first' };
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, message: 'Dedupe: no rows', removed: 0 };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var iEmp = headers.indexOf('Emp ID');
  var iDate = headers.indexOf('Leave Date');
  var iType = headers.indexOf('Leave Type');
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var seen = {};
  var toDelete = [];
  for (var i = 0; i < data.length; i++) {
    var emp = String(data[i][iEmp] || '').trim().toUpperCase();
    var dk = dailyDateKey_(data[i][iDate]);
    var lt = String(data[i][iType] || '').trim();
    if (!emp || !dk) continue;
    var fp = dailyFingerprint_(emp, dk, lt);
    var sheetRow = i + 2;
    if (seen[fp]) toDelete.push(sheetRow);
    else seen[fp] = sheetRow;
  }

  toDelete.sort(function (a, b) { return b - a; });
  var deleted = 0;
  var di = 0;
  while (di < toDelete.length) {
    var blockEnd = toDelete[di];
    var blockStart = blockEnd;
    while (di + 1 < toDelete.length && toDelete[di + 1] === blockStart - 1) {
      di++;
      blockStart = toDelete[di];
    }
    sheet.deleteRows(blockStart, blockEnd - blockStart + 1);
    deleted += (blockEnd - blockStart + 1);
    di++;
  }
  SpreadsheetApp.flush();
  var msg = 'Daily dedupe: removed ' + deleted + ' duplicate day row(s)';
  Logger.log(msg);
  return { success: true, message: msg, removed: deleted };
}

function cleanupDuplicateLeaveRecords() { return runLeaveCleanupPipeline(); }
function cleanupLeaveDuplicates() { return runLeaveCleanupPipeline(); }
