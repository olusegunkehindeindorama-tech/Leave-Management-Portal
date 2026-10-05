/**
 * DiagnoseTblLeave.gs — run diagnoseTblLeave() from the Apps Script editor
 * to see how many real leave rows exist (filters can hide data in the UI).
 */
function diagnoseTblLeave() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tblLeave');
  if (!sh) {
    SpreadsheetApp.getUi().alert('tblLeave sheet not found');
    return { success: false };
  }

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var data = sh.getDataRange().getValues();
  var headers = data.length ? data[0].map(function (h) { return String(h).trim(); }) : [];
  var empIdx = headers.indexOf('Emp ID');
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var yearIdx = headers.indexOf('Entitlement Year');

  var nonEmpty = 0;
  var withDates = 0;
  var yearCounts = {};
  var samples = [];

  for (var i = 1; i < data.length; i++) {
    var emp = empIdx >= 0 ? String(data[i][empIdx] || '').trim() : '';
    if (!emp) continue;
    nonEmpty++;
    var s = startIdx >= 0 ? data[i][startIdx] : '';
    var e = endIdx >= 0 ? data[i][endIdx] : '';
    if (s !== '' && e !== '') withDates++;
    var y = yearIdx >= 0 ? String(data[i][yearIdx] || '') : '';
    if (y) yearCounts[y] = (yearCounts[y] || 0) + 1;
    if (samples.length < 5) {
      samples.push({
        row: i + 1,
        emp: emp,
        start: String(s),
        end: String(e),
        year: y
      });
    }
  }

  var msg =
    'tblLeave diagnosis\n' +
    '-----------------\n' +
    'getLastRow(): ' + lastRow + '\n' +
    'getDataRange() rows: ' + data.length + ' (incl. header)\n' +
    'Rows with Emp ID: ' + nonEmpty + '\n' +
    'Rows with Start+End: ' + withDates + '\n' +
    'Entitlement Year counts: ' + JSON.stringify(yearCounts) + '\n' +
    'Sample rows: ' + JSON.stringify(samples) + '\n\n' +
    'If "Rows with Emp ID" > 0, the sheet is NOT empty.\n' +
    'Check for an active filter on the sheet (Data > Remove filter).';

  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return {
    success: true,
    lastRow: lastRow,
    dataRangeRows: data.length,
    nonEmptyEmp: nonEmpty,
    withDates: withDates,
    yearCounts: yearCounts,
    samples: samples
  };
}

/** Quick import dry-run: counts what Darwinbox import WOULD do without writing. */
function diagnoseDarwinboxImport() {
  var leaveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tblLeave');
  if (!leaveSheet) return { success: false, message: 'tblLeave missing' };
  if (typeof loadLeaveImportContext_ !== 'function') {
    return { success: false, message: 'loadLeaveImportContext_ missing' };
  }
  var ctx = loadLeaveImportContext_(leaveSheet);
  if (!ctx.success) return ctx;
  var csvResult = loadCsvFromFolder_(LEAVE_CSV_FOLDER_ID, DARWINBOX_CSV_NAME);
  if (!csvResult.success) return csvResult;

  var headers = csvResult.headers;
  var rows = csvResult.rows;
  var idx = {
    empId: findHeader_(headers, ['Employee Id', 'Employee ID', 'Emp ID']),
    start: findHeader_(headers, ['Leave From Date', 'From Date', 'Start Date']),
    end: findHeader_(headers, ['Leave To Date', 'To Date', 'End Date']),
    status: findHeader_(headers, ['Status']),
    leaveType: findHeader_(headers, ['Leave Type'])
  };

  var wouldAdd = 0, skippedStatus = 0, skippedDup = 0, skippedBad = 0;
  var keys = {};
  Object.keys(ctx.existingKeys).forEach(function (k) { keys[k] = true; });

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row || !row.length) continue;
    if (idx.status >= 0 && String(row[idx.status] || '').trim() !== 'Approved') {
      skippedStatus++;
      continue;
    }
    var empId = String(row[idx.empId] || '').trim().toUpperCase();
    var startDate = (typeof toCalendarDate_ === 'function' ? toCalendarDate_(row[idx.start]) : parseLeaveDate_(row[idx.start]));
    var endDate = (typeof toCalendarDate_ === 'function' ? toCalendarDate_(row[idx.end]) : parseLeaveDate_(row[idx.end]));
    if (!empId || !startDate || !endDate) { skippedBad++; continue; }
    var fp = fingerprintKey_(empId, startDate, endDate);
    if (keys[fp]) { skippedDup++; continue; }
    keys[fp] = true;
    wouldAdd++;
  }

  var msg = 'Darwinbox dry-run\n' +
    'CSV rows: ' + rows.length + '\n' +
    'Existing fingerprints in tblLeave: ' + Object.keys(ctx.existingKeys).length + '\n' +
    'Would add: ' + wouldAdd + '\n' +
    'Skip dup: ' + skippedDup + '\n' +
    'Skip not-approved: ' + skippedStatus + '\n' +
    'Skip bad dates: ' + skippedBad;
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return {
    success: true,
    message: msg,
    csvRows: rows.length,
    existingKeys: Object.keys(ctx.existingKeys).length,
    wouldAdd: wouldAdd,
    skippedDup: skippedDup,
    skippedStatus: skippedStatus,
    skippedBad: skippedBad
  };
}
