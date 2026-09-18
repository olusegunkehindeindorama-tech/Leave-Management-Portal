/**
 * EMPLOYEE LIST SYNC (trigger-friendly)
 * Source CSV: "Employee List for Leave.csv" in folder 1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1
 *
 * Flow (all in memory, one write):
 *  1. Load CSV + existing tblEmployee
 *  2. Keep / update employees present in CSV
 *  3. Drop employees not in CSV
 *  4. Add new CSV employees
 *  5. Recalculate leave-type columns via batchComputeEmployeeBalances_ (Leave balance.gs)
 *  6. Single clear + setValues back to tblEmployee
 */

var EMP_SYNC_FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
var EMP_SYNC_CSV_NAME = 'Employee List for Leave.csv';

/** Canonical tblEmployee headers (order preserved on write). */
var TBL_EMP_HEADERS = [
  'Business Unit', 'Emp ID', 'Category', 'Status', 'Gender', 'Emp Name',
  'Department', 'Date of Join', 'Date of Release', 'Phone Number', 'Email',
  'Annual', 'Casual', 'Compassionate', 'Examination', 'Study', 'Maternity', 'Probation'
];

/** Leave-type columns on tblEmployee ← full leave type name for balance engine. */
var EMP_LEAVE_COL_MAP = {
  'Annual': 'Annual Leave',
  'Casual': 'Casual Leave',
  'Compassionate': 'Compassionate Leave',
  'Examination': 'Examination Leave',
  'Study': 'Study Leave',
  'Maternity': 'Maternity Leave',
  'Probation': 'Probation Leave'
};

/**
 * Main entry — safe for time-driven triggers and Admin Hub.
 * @returns {{success:boolean, message:string, kept:number, added:number, removed:number}}
 */
function syncEmployeeListFromCsv() {
  var started = new Date().getTime();

  // ---- 1. CSV ----
  var csvResult = loadEmployeeListCsv_();
  if (!csvResult.success) return csvResult;
  var csvMap = csvResult.map; // EmpID -> profile object
  var csvIds = Object.keys(csvMap);

  // ---- 2. Existing tblEmployee into memory ----
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!empSheet) {
    return { success: false, message: 'tblEmployee sheet missing.', kept: 0, added: 0, removed: 0 };
  }

  var existingData = empSheet.getDataRange().getValues();
  var existingHeaders = existingData.length
    ? existingData[0].map(function (h) { return String(h).trim(); })
    : TBL_EMP_HEADERS.slice();

  // Prefer canonical header order; fall back to sheet order if extra cols exist
  var headers = TBL_EMP_HEADERS.slice();
  existingHeaders.forEach(function (h) {
    if (headers.indexOf(h) === -1) headers.push(h);
  });

  var idColName = 'Emp ID';
  var existingMap = {}; // id -> row object keyed by header name
  for (var i = 1; i < existingData.length; i++) {
    var row = existingData[i];
    var obj = {};
    for (var c = 0; c < existingHeaders.length; c++) {
      obj[existingHeaders[c]] = row[c];
    }
    var id = String(obj[idColName] || '').trim().toUpperCase();
    if (id) existingMap[id] = obj;
  }

  // ---- 3. Merge: keep ∩ CSV, add CSV-only, drop sheet-only ----
  var merged = []; // array of row objects
  var kept = 0;
  var added = 0;
  var removed = 0;

  // Keep / update those in CSV
  csvIds.forEach(function (id) {
    var fromCsv = csvMap[id];
    var prev = existingMap[id];
    var rowObj = {};

    // Start from blank canonical fields
    headers.forEach(function (h) { rowObj[h] = ''; });

    // Preserve Date of Release + any prior leave numbers until recalc (if existing)
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

    // Overlay CSV profile fields (source of truth for roster)
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
    // Date of Release: keep previous if any; CSV does not supply it

    merged.push(rowObj);
  });

  // Count removals (were on sheet, not in CSV)
  Object.keys(existingMap).forEach(function (id) {
    if (!csvMap[id]) removed++;
  });

  // ---- 4. Batch load usage + policies + starting bal (once) ----
  var policies = loadPoliciesForBalance_();
  var startingBalMap = loadStartingBalMap_();
  var leaveUsageByEmp = loadLeaveUsageGrouped_(); // empId -> { leaveType -> { year -> util } }

  // ---- 5. Recalculate balances in memory ----
  if (typeof batchComputeEmployeeBalances_ !== 'function') {
    return {
      success: false,
      message: 'batchComputeEmployeeBalances_ not found. Ensure Leave balance.gs is in the project.',
      kept: kept,
      added: added,
      removed: removed
    };
  }

  var today = new Date();
  batchComputeEmployeeBalances_(merged, policies, leaveUsageByEmp, startingBalMap, today, EMP_LEAVE_COL_MAP);

  // ---- 6. Single write ----
  var outRows = [headers];
  for (var m = 0; m < merged.length; m++) {
    var o = merged[m];
    outRows.push(headers.map(function (h) {
      var v = o[h];
      return v === undefined || v === null ? '' : v;
    }));
  }

  // Clear existing data range then write in one shot
  empSheet.clearContents();
  if (outRows.length > 0) {
    empSheet.getRange(1, 1, outRows.length, headers.length).setValues(outRows);
  }

  // Invalidate caches used by the web app
  if (typeof cacheClearAll_ === 'function') cacheClearAll_();
  try {
    if (typeof invalidateEmpCaches_ === 'function') {
      // no per-id list needed after full replace
    }
    CacheService.getScriptCache().remove('emp_map');
  } catch (e) {}

  var ms = new Date().getTime() - started;
  return {
    success: true,
    message: 'Employee sync complete in ' + ms + ' ms. Kept ' + kept +
      ', added ' + added + ', removed ' + removed + '. Balances recalculated for ' + merged.length + ' employees.',
    kept: kept,
    added: added,
    removed: removed,
    total: merged.length,
    elapsedMs: ms
  };
}

// ---------------------------------------------------------------------------
// CSV loader
// ---------------------------------------------------------------------------

function loadEmployeeListCsv_() {
  try {
    var folder = DriveApp.getFolderById(EMP_SYNC_FOLDER_ID);
    var files = folder.getFilesByName(EMP_SYNC_CSV_NAME);
    if (!files.hasNext()) {
      return { success: false, message: 'CSV not found: "' + EMP_SYNC_CSV_NAME + '" in folder ' + EMP_SYNC_FOLDER_ID, map: {} };
    }
    var file = files.next();
    // If multiple same name, take the most recently updated
    while (files.hasNext()) {
      var f2 = files.next();
      if (f2.getLastUpdated() > file.getLastUpdated()) file = f2;
    }

    var text = file.getBlob().getDataAsString();
    var rows = Utilities.parseCsv(text);
    if (!rows || rows.length < 2) {
      return { success: false, message: 'CSV is empty or has no data rows.', map: {} };
    }

    var rawHeaders = rows[0].map(function (h) { return String(h || '').trim(); });
    // Strip "Emp List[...]" wrapper if present
    var headers = rawHeaders.map(function (h) {
      var m = h.match(/^Emp List\[(.+)\]$/i);
      return m ? m[1].trim() : h;
    });

    var idx = function (name) {
      var i = headers.indexOf(name);
      if (i >= 0) return i;
      // soft match
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
  // ISO: 2024-12-09T00:00:00 or 2024-12-09
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  // DD/MM/YYYY
  var dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmy) return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? s : d;
}

// ---------------------------------------------------------------------------
// Shared data loaders (once per run)
// ---------------------------------------------------------------------------

function loadPoliciesForBalance_() {
  if (typeof loadPoliciesCached_ === 'function') return loadPoliciesCached_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sys_LeavePolicies');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var o = {};
    for (var c = 0; c < headers.length; c++) o[headers[c]] = data[i][c];
    rows.push(o);
  }
  return rows;
}

function loadStartingBalMap_() {
  var map = {}; // empId -> number
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('StartingBal');
  if (!sh) return map;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return map;
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var eIdx = headers.indexOf('Emp No');
  if (eIdx < 0) eIdx = headers.indexOf('Emp ID');
  var prevYear = new Date().getFullYear() - 1;
  var bIdx = headers.indexOf(prevYear + ' Balance');
  if (eIdx < 0 || bIdx < 0) return map;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][eIdx] || '').trim().toUpperCase();
    if (id) map[id] = Number(data[i][bIdx]) || 0;
  }
  return map;
}

/**
 * One pass over tblLeave → empId → leaveType → year → sum(utilized)
 */
function loadLeaveUsageGrouped_() {
  var byEmp = {};
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tblLeave');
  if (!sh) return byEmp;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return byEmp;
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var empIdx = headers.indexOf('Emp ID');
  var typeIdx = headers.indexOf('Leave Type');
  var utilIdx = headers.indexOf('Leave Utilized');
  var yearIdx = headers.indexOf('Entitlement Year');
  var startIdx = headers.indexOf('Start Date');
  if (empIdx < 0 || typeIdx < 0) return byEmp;

  var currentYear = new Date().getFullYear();
  for (var i = 1; i < data.length; i++) {
    var empId = String(data[i][empIdx] || '').trim().toUpperCase();
    if (!empId) continue;
    var lt = String(data[i][typeIdx] || '').trim();
    if (!lt) continue;
    var util = utilIdx >= 0 ? (Number(data[i][utilIdx]) || 0) : 0;
    var ey = yearIdx >= 0 && data[i][yearIdx] !== '' && data[i][yearIdx] != null
      ? Number(data[i][yearIdx])
      : (startIdx >= 0 && data[i][startIdx] ? new Date(data[i][startIdx]).getFullYear() : currentYear);
    if (isNaN(ey)) ey = currentYear;

    if (!byEmp[empId]) byEmp[empId] = {};
    if (!byEmp[empId][lt]) byEmp[empId][lt] = {};
    if (!byEmp[empId][lt][ey]) byEmp[empId][lt][ey] = 0;
    byEmp[empId][lt][ey] += util;
  }
  return byEmp;
}
