/**
 * ============================================================
 *  LEAVE RECORD CLEANUP (SAFE — in-place)
 * ============================================================
 *  Dedup key (strict, leave-code ignored):
 *      EmpID | yyyy-MM-dd | yyyy-MM-dd
 *
 *  Keep rule: FIRST occurrence in the sheet (lowest row number).
 *  All later duplicates are deleted — no BP/DB preference.
 *
 *  Entry Code: strip -S1 / -S2 / -a / -b suffixes to base code.
 *  (Recalc later may re-apply -S1/-S2 only where CF split is needed.)
 *
 *  Run:
 *    cleanupDuplicateLeaveRecordsOnly()  — dedupe only
 *    cleanupDuplicateLeaveRecords()      — dedupe then calculateLeaveUtilized
 * ============================================================
 */

function leaveDateKey_(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  try {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    var y = d.getFullYear();
    var m = ('0' + (d.getMonth() + 1)).slice(-2);
    var day = ('0' + d.getDate()).slice(-2);
    return y + '-' + m + '-' + day;
  }
}

function leaveParseDate_(val) {
  if (typeof parseLeaveDate_ === 'function') {
    return parseLeaveDate_(val);
  }
  if (val instanceof Date && !isNaN(val.getTime())) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
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
  var d = new Date(s);
  if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return null;
}

function leaveFingerprint_(empId, startDate, endDate) {
  return String(empId).trim().toUpperCase() + '|' +
    leaveDateKey_(startDate) + '|' + leaveDateKey_(endDate);
}

/**
 * Strip -S1, -S2, -a, -b (and repeats) from Entry Code.
 * BP-11287-a-S1 → BP-11287
 * BP-11073-S2   → BP-11073
 */
function stripEntryCodeSuffix_(code) {
  var c = String(code || '').trim();
  if (!c) return c;
  var prev;
  do {
    prev = c;
    c = c.replace(/-S[12]$/i, '');
    c = c.replace(/-[ab]$/i, '');
  } while (c !== prev);
  return c;
}

/**
 * @param {boolean} optSkipRecalc  true = dedupe only
 */
function cleanupDuplicateLeaveRecords(optSkipRecalc) {
  var started = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) {
    return { success: false, message: 'tblLeave sheet missing.' };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    return { success: true, message: 'No leave rows to clean.', removed: 0 };
  }

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  var empIdx = headers.indexOf('Emp ID');
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var entryIdx = headers.indexOf('Entry Code');

  if (empIdx < 0 || startIdx < 0 || endIdx < 0) {
    return { success: false, message: 'tblLeave missing Emp ID / Start Date / End Date.' };
  }

  // Group by emp|start|end — leave code ignored
  var groups = {}; // fp → [{ sheetRow, dataIdx, startNorm, endNorm }]
  var ungrouped = 0;

  for (var i = 1; i < data.length; i++) {
    var empId = String(data[i][empIdx] || '').trim().toUpperCase();
    var s = leaveParseDate_(data[i][startIdx]) || leaveParseDate_(display[i][startIdx]);
    var e = leaveParseDate_(data[i][endIdx]) || leaveParseDate_(display[i][endIdx]);

    if (!empId || !s || !e) {
      ungrouped++;
      continue;
    }

    var fp = leaveFingerprint_(empId, s, e);
    if (!groups[fp]) groups[fp] = [];
    groups[fp].push({
      sheetRow: i + 1,
      dataIdx: i,
      startNorm: s,
      endNorm: e
    });
  }

  var rowsToDelete = [];
  var rowsToUpdate = [];
  var dupGroups = 0;

  Object.keys(groups).forEach(function (fp) {
    var list = groups[fp];

    // FIRST occurrence only (lowest sheet row) — no BP/DB preference
    list.sort(function (a, b) { return a.sheetRow - b.sheetRow; });
    var keep = list[0];

    var upd = { sheetRow: keep.sheetRow };
    var changed = false;

    // Always strip -S1/-S2/-a/-b from Entry Code on the kept row
    if (entryIdx >= 0) {
      var raw = String(data[keep.dataIdx][entryIdx] || '');
      var clean = stripEntryCodeSuffix_(raw);
      if (clean !== raw) {
        upd.entryCode = clean;
        changed = true;
      }
    }

    // Normalize dates to real Date values + consistent display later
    upd.start = keep.startNorm;
    upd.end = keep.endNorm;
    changed = true;
    if (changed) rowsToUpdate.push(upd);

    if (list.length > 1) {
      dupGroups++;
      for (var d = 1; d < list.length; d++) {
        rowsToDelete.push(list[d].sheetRow);
      }
    }
  });

  // Also strip -S1/-S2 on rows that were never in a parseable group? 
  // Cover ALL data rows for entry-code cleanup (including unique ones already handled above).
  // Unique non-dup rows already in rowsToUpdate. Rows that couldn't parse dates:
  if (entryIdx >= 0) {
    for (var r = 1; r < data.length; r++) {
      var empCheck = String(data[r][empIdx] || '').trim();
      if (!empCheck) continue;
      var rawCode = String(data[r][entryIdx] || '');
      var cleanCode = stripEntryCodeSuffix_(rawCode);
      if (cleanCode === rawCode) continue;
      // Skip if already queued for this sheet row
      var already = false;
      for (var u = 0; u < rowsToUpdate.length; u++) {
        if (rowsToUpdate[u].sheetRow === r + 1) {
          rowsToUpdate[u].entryCode = cleanCode;
          already = true;
          break;
        }
      }
      if (!already) {
        rowsToUpdate.push({ sheetRow: r + 1, entryCode: cleanCode });
      }
    }
  }

  // Apply updates before deletes
  rowsToUpdate.forEach(function (u) {
    if (u.entryCode !== undefined && entryIdx >= 0) {
      sheet.getRange(u.sheetRow, entryIdx + 1).setValue(u.entryCode);
    }
    if (u.start !== undefined) {
      sheet.getRange(u.sheetRow, startIdx + 1).setValue(u.start);
    }
    if (u.end !== undefined) {
      sheet.getRange(u.sheetRow, endIdx + 1).setValue(u.end);
    }
  });

  // Delete later duplicates bottom-up
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

  try {
    var newLast = sheet.getLastRow();
    if (newLast >= 2) {
      sheet.getRange(2, startIdx + 1, newLast - 1, 1).setNumberFormat('dd-mmm-yyyy');
      sheet.getRange(2, endIdx + 1, newLast - 1, 1).setNumberFormat('dd-mmm-yyyy');
    }
  } catch (fmtErr) {}

  if (typeof cacheClearAll_ === 'function') cacheClearAll_();

  var recalcResult = null;
  if (!optSkipRecalc && typeof calculateLeaveUtilized === 'function') {
    try {
      recalcResult = calculateLeaveUtilized();
    } catch (re) {
      recalcResult = { success: false, message: 'Recalc error: ' + re.message };
    }
  }

  var ms = new Date().getTime() - started;
  var msg = 'Cleanup in ' + ms + ' ms: ' + dupGroups + ' duplicate group(s), removed ' +
    deleted + ' later row(s) (kept first occurrence each). Stripped -S1/-S2 from Entry Codes. ' +
    ungrouped + ' row(s) skipped (no key — kept).';
  if (recalcResult && recalcResult.message) msg += ' | Recalc: ' + recalcResult.message;

  Logger.log(msg);
  return {
    success: true,
    message: msg,
    removed: deleted,
    dupGroups: dupGroups,
    updated: rowsToUpdate.length,
    ungroupedKept: ungrouped,
    recalc: recalcResult,
    elapsedMs: ms
  };
}

function cleanupLeaveDuplicates() {
  return cleanupDuplicateLeaveRecords(false);
}

function cleanupDuplicateLeaveRecordsOnly() {
  return cleanupDuplicateLeaveRecords(true);
}
