/**
 * ============================================================
 *  LEAVE IMPORT HELPERS + DARWINBOX IMPORT (trigger-friendly)
 * ============================================================
 *  Dedup on import: EmpID | yyyy-MM-dd | yyyy-MM-dd
 *  After append: scheduleLeaveCleanupPipeline_() (~6 minutes)
 *  Pipeline is NOT run inline (too long for import execution).
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
    var startDate = parseLeaveDate_(row[idx.start]);
    var endDate = parseLeaveDate_(row[idx.end]);
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
    setLeaveCol_(newRow, ctx.lHeaders, 'Start Date', startDate);
    setLeaveCol_(newRow, ctx.lHeaders, 'End Date', endDate);
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

  // Pipeline is heavy — schedule ~6 min later (not inline)
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

function loadLeaveImportContext_(leaveSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');

  if (!empSheet) {
    return { success: false, message: 'tblEmployee sheet missing.' };
  }

  var policyMap = {};
  if (policySheet) {
    var pData = policySheet.getDataRange().getValues();
    if (pData.length > 1) {
      var pH = pData[0].map(function (h) { return String(h).trim(); });
      var pDBName = findHeader_(pH, ['DB Leave Name']);
      var pDBCode = findHeader_(pH, ['DB Leave Code', 'Leave Code']);
      var pStd = findHeader_(pH, ['Leave Type']);
      for (var p = 1; p < pData.length; p++) {
        var dbName = pDBName >= 0 ? String(pData[p][pDBName] || '').trim().toLowerCase() : '';
        if (!dbName) continue;
        policyMap[dbName] = {
          stdType: pStd >= 0 ? String(pData[p][pStd] || '').trim() : '',
          dbCode: pDBCode >= 0 ? String(pData[p][pDBCode] || '').trim() : ''
        };
      }
    }
  }

  var empMap = {};
  var eData = empSheet.getDataRange().getValues();
  if (eData.length > 1) {
    var eH = eData[0].map(function (h) { return String(h).trim(); });
    var eId = findHeader_(eH, ['Emp ID', 'Employee ID']);
    var eName = findHeader_(eH, ['Emp Name', 'Employee Name', 'Name']);
    var eBu = findHeader_(eH, ['Business Unit', 'BU']);
    var eCat = findHeader_(eH, ['Category']);
    var eDept = findHeader_(eH, ['Department', 'Dept']);
    for (var e = 1; e < eData.length; e++) {
      var id = eId >= 0 ? String(eData[e][eId] || '').trim().toUpperCase() : '';
      if (!id) continue;
      empMap[id] = {
        name: eName >= 0 ? String(eData[e][eName] || '').trim() : '',
        bu: eBu >= 0 ? String(eData[e][eBu] || '').trim() : '',
        cat: eCat >= 0 ? String(eData[e][eCat] || '').trim() : '',
        dept: eDept >= 0 ? String(eData[e][eDept] || '').trim() : ''
      };
    }
  }

  var leaveData = leaveSheet.getDataRange().getValues();
  var leaveDisplay = leaveSheet.getDataRange().getDisplayValues();
  var lHeaders = leaveData.length
    ? leaveData[0].map(function (h) { return String(h).trim(); })
    : defaultLeaveHeaders_();

  if (!leaveData.length) {
    leaveSheet.getRange(1, 1, 1, lHeaders.length).setValues([lHeaders]);
  }

  var lEmp = findHeader_(lHeaders, ['Emp ID']);
  var lStart = findHeader_(lHeaders, ['Start Date']);
  var lEnd = findHeader_(lHeaders, ['End Date']);
  var lEntry = findHeader_(lHeaders, ['Entry Code']);

  var existingKeys = {};
  var existingEntryCodes = {};

  for (var r = 1; r < leaveData.length; r++) {
    var emp = lEmp >= 0 ? String(leaveData[r][lEmp] || '').trim().toUpperCase() : '';
    if (!emp) continue;

    var s = lStart >= 0
      ? (parseLeaveDate_(leaveData[r][lStart]) || parseLeaveDate_(leaveDisplay[r][lStart]))
      : null;
    var en = lEnd >= 0
      ? (parseLeaveDate_(leaveData[r][lEnd]) || parseLeaveDate_(leaveDisplay[r][lEnd]))
      : null;

    if (s && en) {
      existingKeys[fingerprintKey_(emp, s, en)] = true;
    }
    if (lEntry >= 0) {
      var code = String(leaveData[r][lEntry] || '').trim().toUpperCase();
      if (code) {
        existingEntryCodes[code] = true;
        var base = code.replace(/-S[12]$/i, '').replace(/-[A-Z]$/i, '');
        if (base) existingEntryCodes[base] = true;
      }
    }
  }

  return {
    success: true,
    empMap: empMap,
    policyMap: policyMap,
    lHeaders: lHeaders,
    existingKeys: existingKeys,
    existingEntryCodes: existingEntryCodes
  };
}

function defaultLeaveHeaders_() {
  return [
    'Entry Code', 'Leave Code', 'Emp ID', 'Emp Name', 'Department', 'Category',
    'Leave Type', 'Start Date', 'End Date', 'Leave Reason', 'No of Days',
    'Leave Utilized', 'Entitlement Year', 'Date Entered', 'Entered By',
    'Date Modified', 'Modified By', 'BU', 'DB Remark', 'Upload Date', 'Uploaded By'
  ];
}

function buildBlankLeaveRow_(headers) {
  var row = [];
  for (var i = 0; i < headers.length; i++) row.push('');
  return row;
}

function setLeaveCol_(row, headers, name, value) {
  var i = headers.indexOf(name);
  if (i >= 0) row[i] = value;
}

function fingerprintKey_(empId, startDate, endDate) {
  return String(empId).trim().toUpperCase() + '|' +
    formatDateKey(startDate) + '|' + formatDateKey(endDate);
}

function formatDateKey(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
  try {
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    var y = dateObj.getFullYear();
    var m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
    var d = ('0' + dateObj.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
}

function findHeader_(headers, names) {
  for (var n = 0; n < names.length; n++) {
    var want = names[n].toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i] || '').trim().toLowerCase() === want) return i;
    }
  }
  return -1;
}

function loadCsvFromFolder_(folderId, fileName) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFilesByName(fileName);
    if (!files.hasNext()) {
      return { success: false, message: 'CSV not found: "' + fileName + '"' };
    }
    var file = files.next();
    while (files.hasNext()) {
      var f2 = files.next();
      if (f2.getLastUpdated() > file.getLastUpdated()) file = f2;
    }
    var parsed = Utilities.parseCsv(file.getBlob().getDataAsString());
    if (!parsed || parsed.length < 2) {
      return { success: false, message: 'CSV empty: ' + fileName };
    }
    var headers = parsed[0].map(function (h) { return String(h || '').trim(); });
    return { success: true, headers: headers, rows: parsed.slice(1), fileName: file.getName() };
  } catch (err) {
    return { success: false, message: 'CSV load error: ' + err.message };
  }
}

function parseLeaveDate_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  if (typeof val === 'number' ||
      (/^\d+(\.\d+)?$/.test(String(val).trim()) && Number(val) > 20000 && Number(val) < 80000)) {
    return excelSerialToDate_(Number(val));
  }

  var s = String(val).trim();
  var months = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
    jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
  };

  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));

  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
  if (mon) {
    var mi = months[mon[2].toLowerCase()];
    if (mi !== undefined) {
      var y = Number(mon[3]);
      if (y < 100) y = y >= 70 ? 1900 + y : 2000 + y;
      return new Date(y, mi, Number(mon[1]));
    }
  }

  var slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (slash) {
    var day = Number(slash[1]);
    var month = Number(slash[2]);
    var year = Number(slash[3]);
    if (year < 100) year = year >= 70 ? 1900 + year : 2000 + year;
    return new Date(year, month - 1, day);
  }

  var d = new Date(s);
  if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return null;
}

function excelSerialToDate_(serial) {
  var epoch = new Date(1899, 11, 30);
  var whole = Math.floor(serial);
  var d = new Date(epoch.getTime() + whole * 86400000);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function appendLeaveRows_(sheet, rows) {
  if (!rows || !rows.length) return;
  var startRow = Math.max(sheet.getLastRow() + 1, 2);
  var cols = rows[0].length;
  var need = startRow + rows.length - 1 - sheet.getMaxRows();
  if (need > 0) sheet.insertRowsAfter(sheet.getMaxRows(), need + 10);

  var CHUNK = 5000;
  for (var i = 0; i < rows.length; i += CHUNK) {
    var part = rows.slice(i, i + CHUNK);
    sheet.getRange(startRow + i, 1, part.length, cols).setValues(part);
  }
}
