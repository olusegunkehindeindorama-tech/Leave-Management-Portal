/**
 * ============================================================
 *  LEAVE RECORD CLEANUP (SAFE — in-place, no data loss)
 * ============================================================
 *  - Does NOT clear or rewrite the whole sheet
 *  - Does NOT drop rows with unparseable dates
 *  - Only removes true duplicates: same Emp ID + Start + End
 *    (when both dates parse successfully)
 *  - Keeps the highest-quality row in each duplicate group
 *  - Strips Entry Code suffixes (-a, -b, -S1, -S2, …) on kept rows
 *  - Normalizes dates only when parse succeeds
 *  - Optionally runs calculateLeaveUtilized afterwards
 *
 *  Run: cleanupDuplicateLeaveRecords()
 * ============================================================
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

function leaveRowQualityScore_(row, headers) {
  var score = 0;
  var entryIdx = headers.indexOf('Entry Code');
  var codeIdx = headers.indexOf('Leave Code');
  var nameIdx = headers.indexOf('Emp Name');
  var utilIdx = headers.indexOf('Leave Utilized');
  var daysIdx = headers.indexOf('No of Days');
  var reasonIdx = headers.indexOf('Leave Reason');
  var byIdx = headers.indexOf('Entered By');

  var entry = entryIdx >= 0 ? String(row[entryIdx] || '') : '';
  var leaveCode = codeIdx >= 0 ? String(row[codeIdx] || '').trim() : '';
  var name = nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '';
  var util = utilIdx >= 0 ? row[utilIdx] : '';
  var days = daysIdx >= 0 ? row[daysIdx] : '';
  var reason = reasonIdx >= 0 ? String(row[reasonIdx] || '').trim() : '';
  var enteredBy = byIdx >= 0 ? String(row[byIdx] || '').trim() : '';

  if (leaveCode && leaveCode.toUpperCase() !== 'NA') score += 10;
  if (name) score += 5;
  if (util !== '' && util !== null && util !== undefined) score += 3;
  if (days !== '' && days !== null && days !== undefined) score += 2;
  if (reason) score += 1;
  if (entry && entry === stripEntryCodeSuffix_(entry)) score += 4;
  if (/^DB-\d+/i.test(entry)) score -= 1;
  if (enteredBy && enteredBy.toLowerCase() !== 'darwinbox') score += 1;
  return score;
}

/**
 * Safe cleanup — in-place only.
 * @param {boolean} optSkipRecalc  true = dedupe only, no entitlement recalc
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
    return { success: true, message: 'No leave rows to clean.', kept: 0, removed: 0 };
  }

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  var empIdx = headers.indexOf('Emp ID');
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var entryIdx = headers.indexOf('Entry Code');

  if (empIdx < 0 || startIdx < 0 || endIdx < 0) {
    return { success: false, message: 'tblLeave missing Emp ID / Start Date / End Date.' };
  }

  // ---- Pass 1: group only rows that have Emp ID + parseable start + end ----
  // Rows that cannot form a fingerprint are NEVER deleted.
  var groups = {}; // fp → [{ sheetRow, dataIdx, score }]
  var ungrouped = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var empId = String(row[empIdx] || '').trim().toUpperCase();
    var s = (typeof parseLeaveDate_ === 'function') ? parseLeaveDate_(row[startIdx]) : null;
    var e = (typeof parseLeaveDate_ === 'function') ? parseLeaveDate_(row[endIdx]) : null;

    if (!empId || !s || !e) {
      ungrouped++;
      continue; // keep row as-is, do not delete
    }

    var fp = (typeof fingerprintKey_ === 'function')
      ? fingerprintKey_(empId, s, e)
      : (empId + '|' + formatDateKey(s) + '|' + formatDateKey(e));

    if (!groups[fp]) groups[fp] = [];
    groups[fp].push({
      sheetRow: i + 1,   // 1-based sheet row
      dataIdx: i,        // index in data[]
      score: leaveRowQualityScore_(row, headers),
      startNorm: s,
      endNorm: e,
      empId: empId
    });
  }

  // ---- Pass 2: decide keep vs delete ----
  var rowsToDelete = []; // sheet row numbers (1-based)
  var rowsToUpdate = []; // { sheetRow, entryCode?, start?, end? }

  Object.keys(groups).forEach(function (fp) {
    var list = groups[fp];
    if (list.length === 1) {
      // Unique — only clean suffix / normalize date if needed
      var only = list[0];
      var upd = { sheetRow: only.sheetRow };
      var changed = false;

      if (entryIdx >= 0) {
        var rawCode = String(data[only.dataIdx][entryIdx] || '');
        var cleanCode = stripEntryCodeSuffix_(rawCode);
        if (cleanCode !== rawCode) {
          upd.entryCode = cleanCode;
          changed = true;
        }
      }
      // Normalize date cells when original wasn't a proper Date
      if (!(data[only.dataIdx][startIdx] instanceof Date)) {
        upd.start = only.startNorm;
        changed = true;
      }
      if (!(data[only.dataIdx][endIdx] instanceof Date)) {
        upd.end = only.endNorm;
        changed = true;
      }
      if (changed) rowsToUpdate.push(upd);
      return;
    }

    // Duplicates: keep best, delete the rest
    list.sort(function (a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return a.sheetRow - b.sheetRow;
    });
    var keep = list[0];
    var keepUpd = { sheetRow: keep.sheetRow };
    var keepChanged = false;

    if (entryIdx >= 0) {
      var kc = stripEntryCodeSuffix_(String(data[keep.dataIdx][entryIdx] || ''));
      if (kc !== String(data[keep.dataIdx][entryIdx] || '')) {
        keepUpd.entryCode = kc;
        keepChanged = true;
      }
    }
    if (!(data[keep.dataIdx][startIdx] instanceof Date)) {
      keepUpd.start = keep.startNorm;
      keepChanged = true;
    }
    if (!(data[keep.dataIdx][endIdx] instanceof Date)) {
      keepUpd.end = keep.endNorm;
      keepChanged = true;
    }
    if (keepChanged) rowsToUpdate.push(keepUpd);

    for (var d = 1; d < list.length; d++) {
      rowsToDelete.push(list[d].sheetRow);
    }
  });

  // ---- Pass 3: apply updates first (before row numbers shift) ----
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

  // ---- Pass 4: delete duplicates bottom-up (preserves other row numbers) ----
  rowsToDelete.sort(function (a, b) { return b - a; });
  // Batch delete contiguous blocks for speed
  var deleted = 0;
  var i = 0;
  while (i < rowsToDelete.length) {
    var end = rowsToDelete[i];   // highest row in this block
    var start = end;
    while (i + 1 < rowsToDelete.length && rowsToDelete[i + 1] === start - 1) {
      i++;
      start = rowsToDelete[i];
    }
    sheet.deleteRows(start, end - start + 1);
    deleted += (end - start + 1);
    i++;
  }

  SpreadsheetApp.flush();

  // Optional date format on Start/End columns (display only — does not delete data)
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
  var msg = 'Safe cleanup in ' + ms + ' ms. Removed ' + deleted +
    ' duplicate row(s), updated ' + rowsToUpdate.length +
    ' kept row(s). Left ' + ungrouped +
    ' row(s) untouched (missing Emp ID or unparseable dates — NOT deleted).';
  if (recalcResult && recalcResult.message) {
    msg += ' | Recalc: ' + recalcResult.message;
  }
  Logger.log(msg);
  return {
    success: true,
    message: msg,
    removed: deleted,
    updated: rowsToUpdate.length,
    ungroupedKept: ungrouped,
    recalc: recalcResult,
    elapsedMs: ms
  };
}

function cleanupLeaveDuplicates() {
  return cleanupDuplicateLeaveRecords(false);
}

/**
 * Dedupe only — no entitlement recalc / split.
 * Prefer this first if you want to inspect results before recalc.
 */
function cleanupDuplicateLeaveRecordsOnly() {
  return cleanupDuplicateLeaveRecords(true);
}
