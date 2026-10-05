/**
 * ============================================================
 *  LEAVE IMPORT HELPERS + DARWINBOX IMPORT (trigger-friendly)
 * ============================================================
 *  Dedup on import: EmpID | yyyy-MM-dd | yyyy-MM-dd
 *  After append: scheduleLeaveCleanupPipeline_() (~6 minutes)
 *  Pipeline is NOT run inline (too long for import execution).
 *  Start/End: DateUtils toCalendarDate_/toSheetDateValue_ (min 2h WAT nudge).
 * ============================================================
 */

var LEAVE_CSV_FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
var DARWINBOX_CSV_NAME = 'Leave_Application.csv';

function importDarwinBoxLeaves() {
  return importDarwinBoxLeaves_();
}

function importDarwinBoxLeaves_() {
  var started = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leaveSheet = ss.getSheetByName('tblLeave');
  if (!leaveSheet) {
    return { success: false, message: 'tblLeave sheet missing.' };
  }

  var ctx = loadLeaveImportContext_(leaveSheet);
  if (!ctx.success) return ctx;

  var csvResult = loadCsvFromFolder_(LEAVE_CSV_FOLDER_ID, DARWINBOX_CSV_NAME);
  if (!csvResult.success) return csvResult;

  var rows = csvResult.rows;
  var headers = csvResult.headers;

  var idx = {
    empId: findHeader_(headers, ['Employee Id', 'Employee ID', 'Emp ID']),
    empName: findHeader_(headers, ['Employee Name', 'Emp Name']),
    start: findHeader_(headers, ['Leave From Date', 'From Date', 'Start Date']),
    end: findHeader_(headers, ['Leave To Date', 'To Date', 'End Date']),
    status: findHeader_(headers, ['Status']),
    applied: findHeader_(headers, ['Applied On', 'Applied Date']),
    leaveType: findHeader_(headers, ['Leave Type']),
    comment: findHeader_(headers, ['Employee Comment', 'Comment', 'Reason'])
  };

  if (idx.empId < 0 || idx.start < 0 || idx.end < 0 || idx.leaveType < 0) {
    return {
      success: false,
      message: 'Darwinbox CSV missing required columns. Found: ' + headers.join(', ')
    };
  }

  var newRows = [];
  var skippedStatus = 0;
  var skippedDup = 0;
  var skippedBad = 0;

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
    if (!empId || !startDate || !endDate) {
      skippedBad++;
      continue;
    }

    var fp = fingerprintKey_(empId, startDate, endDate);
    if (ctx.existingKeys[fp]) {
      skippedDup++;
      continue;
    }

    var empInfo = ctx.empMap[empId] || { bu: '', cat: '', dept: '', name: '' };
    var dbTypeRaw = String(row[idx.leaveType] || '').trim();
    var mapped = ctx.policyMap[dbTypeRaw.toLowerCase()] || {
      stdType: dbTypeRaw,
      dbCode: ''
    };

    var empName = idx.empName >= 0 ? String(row[idx.empName] || '').trim() : '';
    if (!empName) empName = empInfo.name || '';

    var newRow = buildBlankLeaveRow_(ctx.lHeaders);
    setLeaveCol_(newRow, ctx.lHeaders, 'Entry Code', 'DB-' + Date.now() + '-' + i);
    setLeaveCol_(newRow, ctx.lHeaders, 'Leave Code', mapped.dbCode);
    setLeaveCol_(newRow, ctx.lHeaders, 'Emp ID', empId);
    setLeaveCol_(newRow, ctx.lHeaders, 'Emp Name', empName);
    setLeaveCol_(newRow, ctx.lHeaders, 'Department', empInfo.dept);
    setLeaveCol_(newRow, ctx.lHeaders, 'Category', empInfo.cat);
    setLeaveCol_(newRow, ctx.lHeaders, 'Leave Type', mapped.stdType);
    setLeaveCol_(newRow, ctx.lHeaders, 'Start Date', (typeof toSheetDateValue_ === 'function' ? toSheetDateValue_(startDate) : startDate));
    setLeaveCol_(newRow, ctx.lHeaders, 'End Date', (typeof toSheetDateValue_ === 'function' ? toSheetDateValue_(endDate) : endDate));
    setLeaveCol_(newRow, ctx.lHeaders, 'Leave Reason',
      idx.comment >= 0 ? String(row[idx.comment] || '').trim() : '');
    setLeaveCol_(newRow, ctx.lHeaders, 'Date Entered',
      idx.applied >= 0 ? (parseLeaveDate_(row[idx.applied]) || new Date()) : new Date());
    setLeaveCol_(newRow, ctx.lHeaders, 'Entered By', 'Darwinbox');
    setLeaveCol_(newRow, ctx.lHeaders, 'BU', empInfo.bu);
    setLeaveCol_(newRow, ctx.lHeaders, 'DB Remark', 'Original');
    setLeaveCol_(newRow, ctx.lHeaders, 'Upload Date', new Date());
    setLeaveCol_(newRow, ctx.lHeaders, 'Uploaded By', 'Automation');

    newRows.push(newRow);
    ctx.existingKeys[fp] = true;
  }

  if (newRows.length) {
    appendLeaveRows_(leaveSheet, newRows);
  }

  var schedule = null;
  if (typeof scheduleLeaveCleanupPipeline_ === 'function') {
    try {
      schedule = scheduleLeaveCleanupPipeline_();
    } catch (se) {
      schedule = { success: false, message: se.message };
      Logger.log('Could not schedule cleanup pipeline: ' + se.message);
    }
  } else {
    Logger.log('WARNING: scheduleLeaveCleanupPipeline_ not found — load Leave Cleanup.gs');
  }

  var ms = new Date().getTime() - started;
  var msg = 'Darwinbox import: +' + newRows.length + ' new (skipped dup ' +
    skippedDup + ', not-approved ' + skippedStatus + ', bad ' + skippedBad +
    ') in ' + ms + ' ms.';
  if (schedule && schedule.message) msg += ' | ' + schedule.message;
  Logger.log(msg);
  return {
    success: true,
    message: msg,
    added: newRows.length,
    skippedDup: skippedDup,
    skippedStatus: skippedStatus,
    schedule: schedule,
    elapsedMs: ms
  };
}
