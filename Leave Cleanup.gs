/**
 * ============================================================
 *  LEAVE CLEANUP PIPELINE
 * ============================================================
 *  Rules:
 *  - An employee cannot be on two leaves the same calendar day
 *    (Leave Type ignored — Emp ID only).
 *  - Exact duplicate key: EmpID | yyyy-MM-dd | yyyy-MM-dd
 *  - Overlap: keep the shorter range (if equal length → first row),
 *    split the longer into non-overlapping side pieces, drop the
 *    middle that is covered by the shorter. Entry Code gets -a/-b.
 *  - -S1/-S2 are reserved for carry-forward splits (recalc).
 *
 *  Standalone (no import):
 *    runLeaveCleanupPipeline()
 *      → resolve overlaps → exact dedupe → calculateLeaveUtilized
 *
 *  After import:
 *    import already skips exact emp|start|end
 *    then call runLeaveCleanupPipeline() same as above
 * ============================================================
 */

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
  if (typeof parseLeaveDate_ === 'function') return parseLeaveDate_(val);
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

function stripEntryCodeSuffix_(code) {
  var c = String(code || '').trim();
  if (!c) return c;
  var prev;
  do {
    prev = c;
    c = c.replace(/-S[12]$/i, '');
    c = c.replace(/-[a-z]$/i, ''); // -a, -b, -c …
  } while (c !== prev);
  return c;
}

function dayMs_() { return 86400000; }

function addDays_(dateObj, n) {
  var d = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  d.setDate(d.getDate() + n);
  return d;
}

function inclusiveDays_(start, end) {
  return Math.round((end.getTime() - start.getTime()) / dayMs_()) + 1;
}

function rangesOverlap_(s1, e1, s2, e2) {
  return s1.getTime() <= e2.getTime() && s2.getTime() <= e1.getTime();
}

// ---------------------------------------------------------------------------
// Public entry points
// ---------------------------------------------------------------------------

/** Full pipeline: overlaps → exact dedupe → utilization recalc */
function runLeaveCleanupPipeline() {
  var started = new Date().getTime();
  var overlap = resolveOverlappingLeaves_();
  var dedupe = exactDedupeLeaveRecords_();
  var recalc = null;
  if (typeof calculateLeaveUtilized === 'function') {
    try { recalc = calculateLeaveUtilized(); }
    catch (e) { recalc = { success: false, message: e.message }; }
  }
  var ms = new Date().getTime() - started;
  var msg = 'Pipeline done in ' + ms + ' ms. | Overlap: ' + (overlap.message || '') +
    ' | Dedupe: ' + (dedupe.message || '') +
    (recalc && recalc.message ? ' | Recalc: ' + recalc.message : '');
  Logger.log(msg);
  return {
    success: true,
    message: msg,
    overlap: overlap,
    dedupe: dedupe,
    recalc: recalc,
    elapsedMs: ms
  };
}

/** Aliases */
function cleanupDuplicateLeaveRecords() {
  return runLeaveCleanupPipeline();
}
function cleanupLeaveDuplicates() {
  return runLeaveCleanupPipeline();
}
/** Dedupe + overlap only (no recalc) */
function cleanupDuplicateLeaveRecordsOnly() {
  var overlap = resolveOverlappingLeaves_();
  var dedupe = exactDedupeLeaveRecords_();
  return {
    success: true,
    message: 'Overlap + dedupe only. ' + (overlap.message || '') + ' | ' + (dedupe.message || ''),
    overlap: overlap,
    dedupe: dedupe
  };
}

// ---------------------------------------------------------------------------
// Step 1 — Overlap resolve (Emp ID only)
// ---------------------------------------------------------------------------

/**
 * For each employee, while any two leave ranges overlap:
 *   keep the shorter (if equal length → earlier sheet row),
 *   replace the longer with 0–2 side segments (non-overlapping),
 *   tag new segments Entry Code with -a / -b.
 * Runs iteratively until stable. Also collapses identical ranges.
 */
function resolveOverlappingLeaves_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) return { success: false, message: 'tblLeave missing.' };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { success: true, message: 'No rows.', splits: 0, deleted: 0 };

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  var empIdx = headers.indexOf('Emp ID');
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var entryIdx = headers.indexOf('Entry Code');
  if (empIdx < 0 || startIdx < 0 || endIdx < 0) {
    return { success: false, message: 'Missing Emp ID / Start / End columns.' };
  }

  // Working set of leave objects (mutable)
  var leaves = [];
  for (var i = 1; i < data.length; i++) {
    var empId = String(data[i][empIdx] || '').trim().toUpperCase();
    var s = leaveParseDate_(data[i][startIdx]) || leaveParseDate_(display[i][startIdx]);
    var e = leaveParseDate_(data[i][endIdx]) || leaveParseDate_(display[i][endIdx]);
    if (!empId || !s || !e || e < s) continue;
    leaves.push({
      values: data[i].slice(),
      empId: empId,
      start: s,
      end: e,
      origRow: i + 1,
      alive: true
    });
  }

  var splits = 0;
  var deleted = 0;
  var maxPasses = 50;
  var pass = 0;
  var changed = true;

  while (changed && pass < maxPasses) {
    changed = false;
    pass++;

    // Index alive leaves by emp
    var byEmp = {};
    for (var li = 0; li < leaves.length; li++) {
      if (!leaves[li].alive) continue;
      var id = leaves[li].empId;
      if (!byEmp[id]) byEmp[id] = [];
      byEmp[id].push(li);
    }

    Object.keys(byEmp).forEach(function (empId) {
      if (changed) return; // one fix per outer pass for stability
      var idxs = byEmp[empId];
      // Sort by start, then origRow
      idxs.sort(function (a, b) {
        var da = leaves[a].start.getTime() - leaves[b].start.getTime();
        if (da !== 0) return da;
        return leaves[a].origRow - leaves[b].origRow;
      });

      for (var x = 0; x < idxs.length && !changed; x++) {
        for (var y = x + 1; y < idxs.length && !changed; y++) {
          var A = leaves[idxs[x]];
          var B = leaves[idxs[y]];
          if (!A.alive || !B.alive) continue;
          if (!rangesOverlap_(A.start, A.end, B.start, B.end)) continue;

          // Identical range → keep first (lower origRow), drop other
          if (A.start.getTime() === B.start.getTime() && A.end.getTime() === B.end.getTime()) {
            var drop = A.origRow <= B.origRow ? B : A;
            drop.alive = false;
            deleted++;
            changed = true;
            break;
          }

          var lenA = inclusiveDays_(A.start, A.end);
          var lenB = inclusiveDays_(B.start, B.end);
          var small, large;
          if (lenA < lenB) {
            small = A; large = B;
          } else if (lenB < lenA) {
            small = B; large = A;
          } else {
            // Equal length → keep earlier row intact
            if (A.origRow <= B.origRow) { small = A; large = B; }
            else { small = B; large = A; }
          }

          // Build side segments of large outside small
          var segments = [];
          if (large.start.getTime() < small.start.getTime()) {
            segments.push({
              start: large.start,
              end: addDays_(small.start, -1)
            });
          }
          if (large.end.getTime() > small.end.getTime()) {
            segments.push({
              start: addDays_(small.end, 1),
              end: large.end
            });
          }

          // Kill large; add side pieces as new leaves
          large.alive = false;
          deleted++;

          var baseCode = entryIdx >= 0
            ? stripEntryCodeSuffix_(String(large.values[entryIdx] || 'LV'))
            : 'LV';

          for (var si = 0; si < segments.length; si++) {
            var seg = segments[si];
            if (seg.end.getTime() < seg.start.getTime()) continue;
            var newVals = large.values.slice();
            newVals[startIdx] = seg.start;
            newVals[endIdx] = seg.end;
            if (entryIdx >= 0) {
              newVals[entryIdx] = baseCode + '-' + String.fromCharCode(97 + si); // -a, -b
            }
            // Clear utilization fields — recalc will fill
            var utilIdx = headers.indexOf('Leave Utilized');
            var daysIdx = headers.indexOf('No of Days');
            var yearIdx = headers.indexOf('Entitlement Year');
            if (utilIdx >= 0) newVals[utilIdx] = '';
            if (daysIdx >= 0) newVals[daysIdx] = '';
            if (yearIdx >= 0) newVals[yearIdx] = '';

            leaves.push({
              values: newVals,
              empId: large.empId,
              start: seg.start,
              end: seg.end,
              origRow: 999999, // new
              alive: true
            });
            splits++;
          }
          changed = true;
        }
      }
    });
  }

  // Rebuild sheet from alive leaves (preserve header)
  var outRows = [];
  for (var k = 0; k < leaves.length; k++) {
    if (!leaves[k].alive) continue;
    // Ensure date cells are Date objects
    leaves[k].values[startIdx] = leaves[k].start;
    leaves[k].values[endIdx] = leaves[k].end;
    if (entryIdx >= 0) {
      // Strip any leftover -S1/-S2; keep -a/-b we just assigned
      var code = String(leaves[k].values[entryIdx] || '');
      // Only strip S suffixes, preserve single-letter -a/-b
      code = code.replace(/-S[12]$/i, '');
      leaves[k].values[entryIdx] = code;
    }
    outRows.push(leaves[k].values);
  }

  // Sort for readability: Emp ID, Start
  outRows.sort(function (a, b) {
    var ae = String(a[empIdx] || '').toUpperCase();
    var be = String(b[empIdx] || '').toUpperCase();
    if (ae !== be) return ae < be ? -1 : 1;
    return new Date(a[startIdx]) - new Date(b[startIdx]);
  });

  // Write back without clearContents of whole workbook — replace data region
  var out = [headers].concat(outRows);
  // Clear only used data range then write
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
  var need = out.length - sheet.getMaxRows();
  if (need > 0) sheet.insertRowsAfter(sheet.getMaxRows(), need + 10);

  var CHUNK = 4000;
  for (var c = 0; c < out.length; c += CHUNK) {
    var part = out.slice(c, c + CHUNK);
    sheet.getRange(c + 1, 1, part.length, headers.length).setValues(part);
  }
  SpreadsheetApp.flush();

  try {
    if (outRows.length > 0) {
      sheet.getRange(2, startIdx + 1, outRows.length, 1).setNumberFormat('dd-mmm-yyyy');
      sheet.getRange(2, endIdx + 1, outRows.length, 1).setNumberFormat('dd-mmm-yyyy');
    }
  } catch (e) {}

  var msg = 'Overlap resolve: ' + splits + ' segment(s) created, ' +
    deleted + ' overlapping row(s) removed, ' + outRows.length + ' row(s) remain.';
  Logger.log(msg);
  return { success: true, message: msg, splits: splits, deleted: deleted, remaining: outRows.length };
}

// ---------------------------------------------------------------------------
// Step 2 — Exact emp|start|end dedupe (keep first)
// ---------------------------------------------------------------------------

function exactDedupeLeaveRecords_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) return { success: false, message: 'tblLeave missing.' };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { success: true, message: 'No rows.', removed: 0 };

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

  var seen = {}; // fp → first sheet row
  var rowsToDelete = [];
  var rowsToUpdate = [];

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

    var upd = { sheetRow: sheetRow };
    var ch = false;
    if (entryIdx >= 0) {
      var raw = String(data[i][entryIdx] || '');
      // Strip only -S1/-S2 here; keep -a/-b from overlap splits
      var clean = raw.replace(/-S[12]$/i, '');
      // Also strip pure -S1/-S2 if repeated
      while (/-S[12]$/i.test(clean)) clean = clean.replace(/-S[12]$/i, '');
      if (clean !== raw) { upd.entryCode = clean; ch = true; }
    }
    if (!(data[i][startIdx] instanceof Date)) { upd.start = s; ch = true; }
    if (!(data[i][endIdx] instanceof Date)) { upd.end = e; ch = true; }
    if (ch) rowsToUpdate.push(upd);
  }

  rowsToUpdate.forEach(function (u) {
    if (u.entryCode !== undefined) sheet.getRange(u.sheetRow, entryIdx + 1).setValue(u.entryCode);
    if (u.start !== undefined) sheet.getRange(u.sheetRow, startIdx + 1).setValue(u.start);
    if (u.end !== undefined) sheet.getRange(u.sheetRow, endIdx + 1).setValue(u.end);
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

  var msg = 'Exact dedupe: removed ' + deleted + ' duplicate row(s) (kept first).';
  Logger.log(msg);
  return { success: true, message: msg, removed: deleted };
}
