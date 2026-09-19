/**
 * EMPLOYEE LIST SYNC (trigger-friendly)
 * Source CSV: "Employee List for Leave.csv" in folder 1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1
 *
 * Flow:
 *  1. Load CSV + existing tblEmployee
 *  2. Keep / update employees in CSV; drop those not in CSV; add new
 *  3. Single write of roster fields to tblEmployee
 *  4. Call updateAllEmployeeLeaveBalances() (Leave balance.gs) — single source of truth
 */

var EMP_SYNC_FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
var EMP_SYNC_CSV_NAME = 'Employee List for Leave.csv';

var TBL_EMP_HEADERS = [
  'Business Unit', 'Emp ID', 'Category', 'Status', 'Gender', 'Emp Name',
  'Department', 'Date of Join', 'Date of Release', 'Phone Number', 'Email',
  'Annual', 'Casual', 'Compassionate', 'Examination', 'Study', 'Maternity', 'Probation'
];

function syncEmployeeListFromCsv() {
  var started = new Date().getTime();

  var csvResult = loadEmployeeListCsv_();
  if (!csvResult.success) return csvResult;
  var csvMap = csvResult.map;
  var csvIds = Object.keys(csvMap);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!empSheet) {
    return { success: false, message: 'tblEmployee sheet missing.', kept: 0, added: 0, removed: 0 };
  }

  var existingData = empSheet.getDataRange().getValues();
  var existingHeaders = existingData.length
    ? existingData[0].map(function (h) { return String(h).trim(); })
    : TBL_EMP_HEADERS.slice();

  var headers = TBL_EMP_HEADERS.slice();
  existingHeaders.forEach(function (h) {
    if (headers.indexOf(h) === -1) headers.push(h);
  });

  var existingMap = {};
  for (var i = 1; i < existingData.length; i++) {
    var row = existingData[i];
    var obj = {};
    for (var c = 0; c < existingHeaders.length; c++) {
      obj[existingHeaders[c]] = row[c];
    }
    var id = String(obj['Emp ID'] || '').trim().toUpperCase();
    if (id) existingMap[id] = obj;
  }

  var merged = [];
  var kept = 0;
  var added = 0;
  var removed = 0;

  csvIds.forEach(function (id) {
    var fromCsv = csvMap[id];
    var prev = existingMap[id];
    var rowObj = {};
    headers.forEach(function (h) { rowObj[h] = ''; });

    if (prev) {
      kept++;
      headers.forEach(function (h) {
        if (prev[h] !== undefined && prev[h] !== null && prev[h] !== '') {
          rowObj[h] = prev[h];
        }
      });
    } else {
      added++;
    }

    rowObj['Business Unit'] = fromCsv.bu;
    rowObj['Emp ID'] = fromCsv.empId;
    rowObj['Category'] = fromCsv.category;
    rowObj['Status'] = fromCsv.status;
    rowObj['Gender'] = fromCsv.gender;
    rowObj['Emp Name'] = fromCsv.empName;
    rowObj['Department'] = fromCsv.department;
    rowObj['Date of Join'] = fromCsv.dateOfJoin;
    rowObj['Phone Number'] = fromCsv.phone;
    rowObj['Email'] = fromCsv.email;

    merged.push(rowObj);
  });

  Object.keys(existingMap).forEach(function (id) {
    if (!csvMap[id]) removed++;
  });

  // Write roster only (leave columns recalculated next)
  var leaveCols = ['Annual', 'Casual', 'Compassionate', 'Examination', 'Study', 'Maternity', 'Probation'];
  var outRows = [headers];
  for (var m = 0; m < merged.length; m++) {
    var o = merged[m];
    // Clear leave columns — updateAllEmployeeLeaveBalances will fill them
    leaveCols.forEach(function (lc) { o[lc] = ''; });
    outRows.push(headers.map(function (h) {
      var v = o[h];
      return v === undefined || v === null ? '' : v;
    }));
  }

  empSheet.clearContents();
  if (outRows.length > 0) {
    empSheet.getRange(1, 1, outRows.length, headers.length).setValues(outRows);
  }
  SpreadsheetApp.flush();

  // Single source of truth for balances
  var balResult = null;
  if (typeof updateAllEmployeeLeaveBalances === 'function') {
    try {
      balResult = updateAllEmployeeLeaveBalances();
    } catch (be) {
      balResult = { success: false, message: be.message };
    }
  } else {
    balResult = { success: false, message: 'updateAllEmployeeLeaveBalances not found — load Leave balance.gs' };
  }

  try { CacheService.getScriptCache().remove('emp_map'); } catch (e) {}
  if (typeof cacheClearAll_ === 'function') cacheClearAll_();

  var ms = new Date().getTime() - started;
  var msg = 'Employee sync complete in ' + ms + ' ms. Kept ' + kept +
    ', added ' + added + ', removed ' + removed + '.';
  if (balResult && balResult.message) msg += ' | ' + balResult.message;

  return {
    success: true,
    message: msg,
    kept: kept,
    added: added,
    removed: removed,
    total: merged.length,
    balances: balResult,
    elapsedMs: ms
  };
}

function loadEmployeeListCsv_() {
  try {
    var folder = DriveApp.getFolderById(EMP_SYNC_FOLDER_ID);
    var files = folder.getFilesByName(EMP_SYNC_CSV_NAME);
    if (!files.hasNext()) {
      return { success: false, message: 'CSV not found: "' + EMP_SYNC_CSV_NAME + '"', map: {} };
    }
    var file = files.next();
    while (files.hasNext()) {
      var f2 = files.next();
      if (f2.getLastUpdated() > file.getLastUpdated()) file = f2;
    }

    var text = file.getBlob().getDataAsString();
    var rows = Utilities.parseCsv(text);
    if (!rows || rows.length < 2) {
      return { success: false, message: 'CSV is empty.', map: {} };
    }

    var rawHeaders = rows[0].map(function (h) { return String(h || '').trim(); });
    var headers = rawHeaders.map(function (h) {
      var m = h.match(/^Emp List\[(.+)\]$/i);
      return m ? m[1].trim() : h;
    });

    var idx = function (name) {
      var i = headers.indexOf(name);
      if (i >= 0) return i;
      var lower = name.toLowerCase();
      for (var k = 0; k < headers.length; k++) {
        if (String(headers[k]).toLowerCase() === lower) return k;
      }
      return -1;
    };

    var iBU = idx('Business Unit');
    var iId = idx('Emp ID');
    var iCat = idx('Category');
    var iSt = idx('Status');
    var iGen = idx('Gender');
    var iName = idx('Emp Name');
    var iDept = idx('Department');
    var iDoj = idx('Date of Join');
    var iPhone = idx('Phone Number');
    var iEmail = idx('Email');

    if (iId < 0) {
      return { success: false, message: 'CSV missing Emp ID column.', map: {} };
    }

    var map = {};
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var empId = String(row[iId] || '').trim().toUpperCase();
      if (!empId) continue;

      map[empId] = {
        empId: empId,
        bu: iBU >= 0 ? String(row[iBU] || '').trim() : '',
        category: iCat >= 0 ? String(row[iCat] || '').trim() : '',
        status: iSt >= 0 ? String(row[iSt] || '').trim() : '',
        gender: iGen >= 0 ? String(row[iGen] || '').trim() : '',
        empName: iName >= 0 ? String(row[iName] || '').trim() : '',
        department: iDept >= 0 ? String(row[iDept] || '').trim() : '',
        dateOfJoin: iDoj >= 0 ? parseFlexibleDate_(row[iDoj]) : '',
        phone: iPhone >= 0 ? String(row[iPhone] || '').trim() : '',
        email: iEmail >= 0 ? String(row[iEmail] || '').trim() : ''
      };
    }

    return { success: true, map: map, count: Object.keys(map).length };
  } catch (err) {
    return { success: false, message: 'CSV load error: ' + err.message, map: {} };
  }
}

function parseFlexibleDate_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  var s = String(val).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  var dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmy) return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? s : d;
}
