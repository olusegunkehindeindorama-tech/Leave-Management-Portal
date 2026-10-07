/**
 * DailyCleanup.gs — ONLY leave cleanup for daily tblLeave model.
 * Use these (do NOT use old Leave Cleanup.gs):
 *   runLeaveCleanupPipeline()
 *   scheduleLeaveCleanupPipeline_()
 * Dedup: Emp ID | Leave Date
 */
var CLEANUP_TRIGGER_HANDLER = 'runLeaveCleanupPipelineFromTrigger';

function scheduleLeaveCleanupPipeline_() {
  try {
    deleteTriggersByHandler_(CLEANUP_TRIGGER_HANDLER);
    ScriptApp.newTrigger(CLEANUP_TRIGGER_HANDLER)
      .timeBased()
      .after(6 * 60 * 1000)
      .create();
    Logger.log('scheduleLeaveCleanupPipeline_: daily cleanup in ~6 minutes');
    return { success: true, message: 'Daily cleanup scheduled in ~6 minutes' };
  } catch (e) {
    Logger.log('scheduleLeaveCleanupPipeline_ failed: ' + e.message);
    return { success: false, message: e.message };
  }
}

function runLeaveCleanupPipelineFromTrigger() {
  var result;
  try {
    result = runLeaveCleanupPipeline();
  } catch (e) {
    result = { success: false, message: String(e.message || e) };
    Logger.log('runLeaveCleanupPipelineFromTrigger error: ' + e.message);
  }
  try { deleteTriggersByHandler_(CLEANUP_TRIGGER_HANDLER); } catch (e2) {}
  return result;
}

function deleteTriggersByHandler_(handlerName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function runLeaveCleanupPipeline() {
  Logger.log('=== DAILY PIPELINE START ===');
  var t0 = new Date().getTime();
  var dedupe = exactDedupeDailyLeaveInPlace_();
  Logger.log('[1] dedupe: ' + (dedupe.message || ''));
  var recalc = null;
  if (typeof calculateLeaveUtilizedDaily_ === 'function') {
    recalc = calculateLeaveUtilizedDaily_();
  } else if (typeof calculateLeaveUtilized === 'function') {
    recalc = calculateLeaveUtilized();
  }
  Logger.log('[2] recalc: ' + (recalc && recalc.message ? recalc.message : ''));
  var ms = new Date().getTime() - t0;
  var msg = 'Daily pipeline ' + ms + ' ms | ' + (dedupe.message || '') +
    (recalc && recalc.message ? ' | ' + recalc.message : '');
  Logger.log('=== DAILY PIPELINE END === ' + msg);
  return { success: true, message: msg, dedupe: dedupe, recalc: recalc, elapsedMs: ms };
}

function exactDedupeDailyLeaveInPlace_() {
  Logger.log('exactDedupeDailyLeaveInPlace_: Emp|Date only');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) return { success: false, message: 'tblLeave missing' };
  if (typeof isDailyTblLeaveSchema_ === 'function' && !isDailyTblLeaveSchema_(sheet)) {
    return { success: false, message: 'Not daily schema — run migrateTblLeaveToDaily() first. Do not use old Leave Cleanup.' };
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, message: 'Dedupe: no rows', removed: 0 };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var iEmp = headers.indexOf('Emp ID');
  var iDate = headers.indexOf('Leave Date');
  if (iEmp < 0 || iDate < 0) {
    return { success: false, message: 'Need Emp ID + Leave Date columns' };
  }
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var seen = {};
  var toDelete = [];
  for (var i = 0; i < data.length; i++) {
    var emp = String(data[i][iEmp] || '').trim().toUpperCase();
    var dk = (typeof dailyDateKey_ === 'function') ? dailyDateKey_(data[i][iDate]) : '';
    if (!emp || !dk) continue;
    var fp = (typeof dailyFingerprint_ === 'function') ? dailyFingerprint_(emp, dk) : (emp + '|' + dk);
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
function cleanupDuplicateLeaveRecordsOnly() { return exactDedupeDailyLeaveInPlace_(); }
