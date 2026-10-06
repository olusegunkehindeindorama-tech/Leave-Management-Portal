/**
 * ============================================================
 *  LEAVE CLEANUP PIPELINE (SAFE — in-place only)
 * ============================================================
 *  NEVER clearContents / rewrite the whole sheet.
 *  ONLY: update cells, delete full duplicate/overlap rows, append splits.
 *
 *  Emp ID only for overlaps (leave type ignored).
 *  Exact key: EmpID | yyyy-MM-dd | yyyy-MM-dd
 *
 *  Manual:
 *    runLeaveCleanupPipeline()
 *
 *  After import (DB / Excel):
 *    scheduleLeaveCleanupPipeline_()  → runs in ~6 minutes via trigger
 *    runLeaveCleanupPipelineFromTrigger() → pipeline + delete own trigger
 *      and any other triggers for this handler
 *
 *  Dates: all Start/End writes use leaveSheetDate_ (local noon) so WAT
 *  does not display previous-day 23:00. Prefer DateUtils toCalendarDate_.
 * ============================================================
 */

var CLEANUP_TRIGGER_HANDLER = 'runLeaveCleanupPipelineFromTrigger';
var CLEANUP_TRIGGER_DELAY_MS = 6 * 60 * 1000; // 6 minutes
var CLEANUP_TRIGGER_PROP = 'PENDING_CLEANUP_TRIGGER_ID';

function leaveDateKey_(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  try {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return d.getFullYear() + '-' +
      ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
      ('0' + d.getDate()).slice(-2);
  }
}

function leaveParseDate_(val) {
  // Prefer DateUtils calendar normalize (2h evening nudge for WAT)
  if (typeof toCalendarDate_ === 'function') {
    var cd = toCalendarDate_(val);
    if (cd) return cd;
  }
  if (typeof parseLeaveDate_ === 'function') {
    var p = parseLeaveDate_(val);
    if (p) return p;
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
}

/** Always write Start/End as local noon + dd-mmm-yyyy (no WAT day-back shift). */
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
  if (row % 500 === 0 || row < 5) {
    try {
      Logger.log('leaveWriteDateCell_ r' + row + 'c' + col + ' → ' +
        Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
    } catch (e) {}
  }
}

function leaveFingerprint_(empId, startDate, endDate) {
  return String(empId).trim().toUpperCase() + '|' +
    leaveDateKey_(startDate) + '|' + leaveDateKey_(endDate);
}

function stripSSuffix_(code) {
  var c = String(code || '').trim();
  while (/-S[12]$/i.test(c)) c = c.replace(/-S[12]$/i, '');
  return c;
}

function stripAllSuffixes_(code) {
  var c = String(code || '').trim();
  var prev;
  do {
    prev = c;
    c = c.replace(/-S[12]$/i, '');
    c = c.replace(/-[a-z]$/i, '');
  } while (c !== prev);
  return c;
}

function addDays_(dateObj, n) {
  var d = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  d.setDate(d.getDate() + n);
  return d;
}

function inclusiveDays_(start, end) {
  return Math.round((end.getTime() - start.getTime()) / 86400000) + 1;
}

function rangesOverlap_(s1, e1, s2, e2) {
  return s1.getTime() <= e2.getTime() && s2.getTime() <= e1.getTime();
}

function scheduleLeaveCleanupPipeline_() {
  Logger.log('scheduleLeaveCleanupPipeline_: scheduling cleanup in ~6 minutes');
  deleteTriggersByHandler_(CLEANUP_TRIGGER_HANDLER);
  var trigger = ScriptApp.newTrigger(CLEANUP_TRIGGER_HANDLER)
    .timeBased()
    .after(CLEANUP_TRIGGER_DELAY_MS)
    .create();
  try {
    PropertiesService.getScriptProperties().setProperty(
      CLEANUP_TRIGGER_PROP,
      trigger.getUniqueId()
    );
  } catch (e) {}
  Logger.log('Scheduled ' + CLEANUP_TRIGGER_HANDLER + ' in ~6 minutes (id=' +
    trigger.getUniqueId() + ')');
  return {
    success: true,
    message: 'Cleanup pipeline scheduled in ~6 minutes.',
    triggerId: trigger.getUniqueId()
  };
}

function runLeaveCleanupPipelineFromTrigger() {
  Logger.log('runLeaveCleanupPipelineFromTrigger: begin');
  var result = null;
  try {
    result = runLeaveCleanupPipeline();
  } catch (e) {
    result = { success: false, message: e.message };
    Logger.log('Pipeline from trigger failed: ' + e.message);
  } finally {
    try {
      cleanupImportRelatedTriggers_();
    } catch (te) {
      Logger.log('Trigger cleanup error: ' + te.message);
    }
  }
  return result;
}

function cleanupImportRelatedTriggers_() {
  var deleted = deleteTriggersByHandler_(CLEANUP_TRIGGER_HANDLER);
  try {
    PropertiesService.getScriptProperties().deleteProperty(CLEANUP_TRIGGER_PROP);
  } catch (e) {}
  Logger.log('Deleted ' + deleted + ' trigger(s) for ' + CLEANUP_TRIGGER_HANDLER);
  return deleted;
}

function deleteTriggersByHandler_(handlerName) {
  var count = 0;
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    try {
      if (triggers[i].getHandlerFunction() === handlerName) {
        ScriptApp.deleteTrigger(triggers[i]);
        count++;
      }
    } catch (e) {
      Logger.log('Could not delete trigger: ' + e.message);
    }
  }
  return count;
}

function runLeaveCleanupPipeline() {
  var started = new Date().getTime();
  Logger.log('=== PIPELINE START ===');

  Logger.log('[1/3] resolveOverlappingLeavesInPlace_ ...');
  var overlap = resolveOverlappingLeavesInPlace_();
  Logger.log('[1/3] done: ' + (overlap && overlap.message ? overlap.message : JSON.stringify(overlap)));

  Logger.log('[2/3] exactDedupeLeaveRecordsInPlace_ (noon-normalize Start/End) ...');
  var dedupe = exactDedupeLeaveRecordsInPlace_();
  Logger.log('[2/3] done: ' + (dedupe && dedupe.message ? dedupe.message : JSON.stringify(dedupe)));

  var recalc = null;
  Logger.log('[3/3] calculateLeaveUtilized / recalculate ...');
  if (typeof calculateLeaveUtilized === 'function') {
    try {
      recalc = calculateLeaveUtilized();
      Logger.log('[3/3] done: ' + (recalc && recalc.message ? recalc.message : JSON.stringify(recalc)));
    } catch (e) {
      recalc = { success: false, message: e.message };
      Logger.log('[3/3] FAILED: ' + e.message);
    }
  } else {
    Logger.log('[3/3] skipped — calculateLeaveUtilized not found');
  }

  var ms = new Date().getTime() - started;
  var msg = 'Pipeline ' + ms + ' ms. | ' + (overlap.message || '') +
    ' | ' + (dedupe.message || '') +
    (recalc && recalc.message ? ' | ' + recalc.message : '');
  Logger.log('=== PIPELINE END === ' + msg);
  return {
    success: true,
    message: msg,
    overlap: overlap,
    dedupe: dedupe,
    recalc: recalc,
    elapsedMs: ms
  };
}

function cleanupDuplicateLeaveRecords() { return runLeaveCleanupPipeline(); }
function cleanupLeaveDuplicates() { return runLeaveCleanupPipeline(); }
function cleanupDuplicateLeaveRecordsOnly() {
  Logger.log('cleanupDuplicateLeaveRecordsOnly: begin');
  var overlap = resolveOverlappingLeavesInPlace_();
  var dedupe = exactDedupeLeaveRecordsInPlace_();
  return {
    success: true,
    message: (overlap.message || '') + ' | ' + (dedupe.message || ''),
    overlap: overlap,
    dedupe: dedupe
  };
}

function resolveOverlappingLeavesInPlace_() {
  Logger.log('resolveOverlappingLeavesInPlace_: begin');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) return { success: false, message: 'tblLeave missing.' };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { success: true, message: 'Overlap: no rows.', splits: 0, deleted: 0 };

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });
  var empIdx = headers.indexOf('Emp ID');
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var entryIdx = headers.indexOf('Entry Code');
  if (empIdx < 0 || startIdx < 0 || endIdx < 0) {
    return { success: false, message: 'Missing Emp ID / Start / End.' };
  }

  var totalSplits = 0;
  var totalDeleted = 0;
  var maxOuter = 30;

  for (var outer = 0; outer < maxOuter; outer++) {
    lastRow = sheet.getLastRow();
    if (lastRow < 2) break;

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

    var leaves = [];
    for (var i = 1; i < data.length; i++) {
      var empId = String(data[i][empIdx] || '').trim().toUpperCase();
      var s = leaveParseDate_(data[i][startIdx]) || leaveParseDate_(display[i][startIdx]);
      var e = leaveParseDate_(data[i][endIdx]) || leaveParseDate_(display[i][endIdx]);
      if (!empId || !s || !e || e < s) continue;
      leaves.push({
        row: i + 1,
        empId: empId,
        start: s,
        end: e,
        values: data[i].slice()
      });
    }

    var byEmp = {};
    for (var li = 0; li < leaves.length; li++) {
      var id = leaves[li].empId;
      if (!byEmp[id]) byEmp[id] = [];
      byEmp[id].push(leaves[li]);
    }

    var fixedOne = false;

    Object.keys(byEmp).forEach(function (empId) {
      if (fixedOne) return;
      var list = byEmp[empId];
      list.sort(function (a, b) {
        var d = a.start.getTime() - b.start.getTime();
        if (d !== 0) return d;
        return a.row - b.row;
      });

      for (var x = 0; x < list.length && !fixedOne; x++) {
        for (var y = x + 1; y < list.length && !fixedOne; y++) {
          var A = list[x];
          var B = list[y];
          if (!rangesOverlap_(A.start, A.end, B.start, B.end)) continue;

          if (A.start.getTime() === B.start.getTime() && A.end.getTime() === B.end.getTime()) {
            var dropRow = A.row < B.row ? B.row : A.row;
            sheet.deleteRow(dropRow);
            totalDeleted++;
            fixedOne = true;
            break;
          }

          var lenA = inclusiveDays_(A.start, A.end);
          var lenB = inclusiveDays_(B.start, B.end);
          var small, large;
          if (lenA < lenB) { small = A; large = B; }
          else if (lenB < lenA) { small = B; large = A; }
          else {
            if (A.row <= B.row) { small = A; large = B; }
            else { small = B; large = A; }
          }

          var segments = [];
          if (large.start.getTime() < small.start.getTime()) {
            var leftEnd = addDays_(small.start, -1);
            if (leftEnd.getTime() >= large.start.getTime()) {
              segments.push({ start: large.start, end: leftEnd });
            }
          }
          if (large.end.getTime() > small.end.getTime()) {
            var rightStart = addDays_(small.end, 1);
            if (rightStart.getTime() <= large.end.getTime()) {
              segments.push({ start: rightStart, end: large.end });
            }
          }

          var baseCode = entryIdx >= 0
            ? stripAllSuffixes_(String(large.values[entryIdx] || 'LV'))
            : 'LV';

          var appendRows = [];
          for (var si = 0; si < segments.length; si++) {
            var seg = segments[si];
            var newVals = large.values.slice();
            while (newVals.length < lastCol) newVals.push('');
            newVals[startIdx] = leaveSheetDate_(seg.start);
            newVals[endIdx] = leaveSheetDate_(seg.end);
            Logger.log('Overlap split segment ' + si + ' dates → noon');
            if (entryIdx >= 0) {
              newVals[entryIdx] = baseCode + '-' + String.fromCharCode(97 + si);
            }
            var utilIdx = headers.indexOf('Leave Utilized');
            var daysIdx = headers.indexOf('No of Days');
            var yearIdx = headers.indexOf('Entitlement Year');
            if (utilIdx >= 0) newVals[utilIdx] = '';
            if (daysIdx >= 0) newVals[daysIdx] = '';
            if (yearIdx >= 0) newVals[yearIdx] = '';
            appendRows.push(newVals);
          }

          sheet.deleteRow(large.row);
          totalDeleted++;

          if (appendRows.length) {
            if (typeof appendLeaveRows_ === 'function') {
              appendLeaveRows_(sheet, appendRows);
            } else {
              appendRows.forEach(function (r) { sheet.appendRow(r); });
            }
            totalSplits += appendRows.length;
          }

          SpreadsheetApp.flush();
          fixedOne = true;
        }
      }
    });

    if (!fixedOne) break;
  }

  var msg = 'Overlap: deleted ' + totalDeleted + ' row(s), appended ' +
    totalSplits + ' segment(s).';
  Logger.log(msg);
  return { success: true, message: msg, deleted: totalDeleted, splits: totalSplits };
}

function exactDedupeLeaveRecordsInPlace_() {
  Logger.log('exactDedupeLeaveRecordsInPlace_: begin (will noon-normalize dates)');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) return { success: false, message: 'tblLeave missing.' };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { success: true, message: 'Dedupe: no rows.', removed: 0 };

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  var empIdx = headers.indexOf('Emp ID');
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var entryIdx = headers.indexOf('Entry Code');
  if (empIdx < 0 || startIdx < 0 || endIdx < 0) {
    return { success: false, message: 'Missing columns.' };
  }

  var seen = {};
  var rowsToDelete = [];
  var entryUpdates = [];

  for (var i = 1; i < data.length; i++) {
    var empId = String(data[i][empIdx] || '').trim().toUpperCase();
    var s = leaveParseDate_(data[i][startIdx]) || leaveParseDate_(display[i][startIdx]);
    var e = leaveParseDate_(data[i][endIdx]) || leaveParseDate_(display[i][endIdx]);
    if (!empId || !s || !e) continue;

    var fp = leaveFingerprint_(empId, s, e);
    var sheetRow = i + 1;

    if (seen[fp]) {
      rowsToDelete.push(sheetRow);
      continue;
    }
    seen[fp] = sheetRow;

    if (entryIdx >= 0) {
      var raw = String(data[i][entryIdx] || '');
      var clean = stripSSuffix_(raw);
      if (clean !== raw) {
        entryUpdates.push({ row: sheetRow, code: clean });
      }
    }

    // Normalize Start/End to noon calendar dates (fixes 23:00 previous-day artifacts)
    if (s) leaveWriteDateCell_(sheet, sheetRow, startIdx + 1, s);
    if (e) leaveWriteDateCell_(sheet, sheetRow, endIdx + 1, e);
  }

  entryUpdates.forEach(function (u) {
    sheet.getRange(u.row, entryIdx + 1).setValue(u.code);
  });

  rowsToDelete.sort(function (a, b) { return b - a; });
  var deleted = 0;
  var di = 0;
  while (di < rowsToDelete.length) {
    var blockEnd = rowsToDelete[di];
    var blockStart = blockEnd;
    while (di + 1 < rowsToDelete.length && rowsToDelete[di + 1] === blockStart - 1) {
      di++;
      blockStart = rowsToDelete[di];
    }
    sheet.deleteRows(blockStart, blockEnd - blockStart + 1);
    deleted += (blockEnd - blockStart + 1);
    di++;
  }
  SpreadsheetApp.flush();

  var msg = 'Dedupe: removed ' + deleted + ' full duplicate row(s).';
  Logger.log(msg);
  return { success: true, message: msg, removed: deleted };
}
