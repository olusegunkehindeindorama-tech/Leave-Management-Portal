
/**
 * ============================================================
 *  SHIFT DATA PROCESSOR  — Google Apps Script
 * ============================================================
 *  WHAT IT DOES
 *  - Reads "YYYY Month.xlsx" files from SOURCE_FOLDER_ID
 *  - Extracts Emp No + daily shift codes, unpivots to {Emp ID, Date, Shift}
 *  - INCREMENTAL: only re-processes files modified since the last run
 *  - Destination sheet is updated in-place (unchanged months kept as-is)
 *  - Scheduled to run daily at 1 AM Nigerian time (WAT = UTC+1)
 *

 *
 *  TO FORCE A FULL REFRESH  (e.g. after manual file edits):
 *    Run  clearStateAndReprocess()  from the editor.
 * ============================================================
 */

// ─── CONFIGURATION ──────────────────────────────────────────
var SOURCE_FOLDER_ID = '1WX2RK3b5lPegwNrLzVrWs7vJ-_sWQnF1';
var DEST_FILE_ID     = '13ahfJJLgh3lBO4VPSWG6nRau_9MIPrryfNrxPs4veKA';
var DEST_SHEET_NAME  = 'tblShift';   // ← exact tab name in destination file
var QUEUE_KEY = 'shiftProcessingQueue';
var WORKER_TRIGGER = 'processShiftFiles';
// ────────────────────────────────────────────────────────────

// Columns to ignore when scanning source header rows
var DISCARD_HEADERS = ['name', 'work group', 'shift/ general', 'shift/general',
                       'function dept', 'dept'];

// PropertiesService key where we store per-file modification timestamps
var PROP_KEY = 'shiftFilesMeta';


// ============================================================
//  MAIN ENTRY POINT
// ============================================================
function processShiftFiles() {

  var props = PropertiesService.getScriptProperties();

  var queue = JSON.parse(props.getProperty(QUEUE_KEY) || '[]');

  //
  // FIRST RUN OF A CYCLE
  // Build queue only if queue is empty
  //
  if (queue.length === 0) {

    Logger.log('Building queue...');

    var storedMeta = JSON.parse(props.getProperty(PROP_KEY) || '{}');

    var folder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    var iter = folder.getFiles();

    while (iter.hasNext()) {

      var file = iter.next();

      var fileName = file.getName();

      var match = fileName.match(/^(\d{4})\s+([A-Za-z]+)\.xlsx$/i);

      if (!match) continue;

      var year = parseInt(match[1], 10);

      var monthIndex = parseMonthName(match[2]);

      if (monthIndex === -1) continue;

      var fileId = file.getId();

      var modified = file.getLastUpdated().getTime();

      var stored = storedMeta[fileId];

      if (!stored || stored.modified !== modified) {

        queue.push({
          fileId: fileId,
          fileName: fileName,
          year: year,
          monthIndex: monthIndex,
          modified: modified,
          isModified: !!stored
        });

      }
    }

    props.setProperty(QUEUE_KEY, JSON.stringify(queue));

    Logger.log('Queue built. Files to process: ' + queue.length);
  }

  //
  // NOTHING TO DO
  //
  if (queue.length === 0) {

    Logger.log('No new or modified files.');

    deleteWorkerTriggers_();

    return;
  }

  //
  // PROCESS EXACTLY ONE FILE
  //
  var entry = queue.shift();

  props.setProperty(QUEUE_KEY, JSON.stringify(queue));

  Logger.log('Processing: ' + entry.fileName);

  processSingleFile_(entry);

  //
  // CONTINUE IF MORE FILES EXIST
  //
  if (queue.length > 0) {

    scheduleWorker_();

    Logger.log('Remaining files: ' + queue.length);

  } else {

    deleteWorkerTriggers_();

    Logger.log('Queue completed.');
  }
}

function processSingleFile_(entry) {

  var props = PropertiesService.getScriptProperties();

  var storedMeta = JSON.parse(
    props.getProperty(PROP_KEY) || '{}'
  );

  var destSS = SpreadsheetApp.openById(DEST_FILE_ID);

  var destSheet = destSS.getSheetByName(DEST_SHEET_NAME);

  if (!destSheet) {
    destSheet = destSS.insertSheet(DEST_SHEET_NAME);
  }

  var tempId = null;

  try {

    //
    // DELETE OLD RECORDS IF FILE WAS MODIFIED
    //
    if (entry.isModified) {

      Logger.log('Removing existing rows...');

      removeMonthRows_(
        destSheet,
        entry.year,
        entry.monthIndex
      );
    }

    //
    // CONVERT
    //
    tempId = convertXlsxToGSheet(
      DriveApp.getFileById(entry.fileId)
    );

    //
    // EXTRACT
    //
    var rows = extractShiftRows(
      tempId,
      entry.year,
      entry.monthIndex
    );

    //
    // WRITE
    //
    writeChunked_(destSheet, rows);

    //
    // CHECKPOINT
    //
    storedMeta[entry.fileId] = {
      modified: entry.modified,
      year: entry.year,
      month: entry.monthIndex
    };

    props.setProperty(
      PROP_KEY,
      JSON.stringify(storedMeta)
    );

    Logger.log(
      'Finished: ' +
      entry.fileName +
      ' (' +
      rows.length +
      ' rows)'
    );

  } finally {

    if (tempId) {

      try {
        DriveApp.getFileById(tempId)
          .setTrashed(true);
      } catch(e) {}

    }
  }
}

function removeMonthRows_(sheet, year, monthIndex) {

  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) return;

  var kept = [];

  for (var r = 1; r < data.length; r++) {

    var row = data[r];

    var dateVal = row[1];

    var ry, rm;

    if (dateVal instanceof Date) {

      ry = dateVal.getFullYear();
      rm = dateVal.getMonth();

    } else {

      var parts = String(dateVal).split('-');

      ry = parseInt(parts[0], 10);
      rm = parseInt(parts[1], 10) - 1;
    }

    if (!(ry === year && rm === monthIndex)) {

      if (dateVal instanceof Date) {
        row[1] = Utilities.formatDate(
          dateVal,
          'Africa/Lagos',
          'yyyy-MM-dd'
        );
      }

      kept.push(row);
    }
  }

  sheet.clearContents();

  sheet.getRange(1, 1, 1, 3)
    .setValues([['Emp ID','Date','Shift']]);

  writeChunked_(sheet, kept);
}



function scheduleWorker_() {

  ScriptApp.newTrigger(WORKER_TRIGGER)
    .timeBased()
    .after(1000)
    .create();
}


function deleteWorkerTriggers_() {

  var triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(function(t){

    if (
      t.getHandlerFunction() ===
      WORKER_TRIGGER
    ) {
      ScriptApp.deleteTrigger(t);
    }

  });
}


// ============================================================
//  FORCE FULL REFRESH  (run manually when needed)
// ============================================================
/**
 * Clears stored modification metadata so the next run of
 * processShiftFiles() re-processes every file from scratch.
 */
function clearStateAndReprocess() {

  var props = PropertiesService.getScriptProperties();

  props.deleteProperty(PROP_KEY);
  props.deleteProperty(QUEUE_KEY);

  deleteWorkerTriggers_();

  var destSS = SpreadsheetApp.openById(DEST_FILE_ID);
  var destSheet = destSS.getSheetByName(DEST_SHEET_NAME);

  if (destSheet) {
    destSheet.clearContents();
    destSheet.getRange(1,1,1,3)
      .setValues([['Emp ID','Date','Shift']]);
  }

  processShiftFiles();
}


function writeChunked_(sheet, rows) {

  if (!rows || rows.length === 0) return;

  var CHUNK_SIZE = 10000;

  var startRow = sheet.getLastRow() + 1;

  for (var i = 0; i < rows.length; i += CHUNK_SIZE) {

    var chunk = rows.slice(i, i + CHUNK_SIZE);

    sheet
      .getRange(startRow, 1, chunk.length, 3)
      .setValues(chunk);

    startRow += chunk.length;
  }
}

// ============================================================
//  HELPER: Convert .xlsx → temporary Google Sheet (3-tier fallback)
// ============================================================
function convertXlsxToGSheet(file) {
  var blob     = file.getBlob();
  var tempName = '_TEMP_SHIFT_' + file.getId();

  // Method 1: Drive API v3
  try {
    if (typeof Drive !== 'undefined' && Drive.Files &&
        typeof Drive.Files.create === 'function') {
      var r3 = Drive.Files.create(
        { name: tempName, mimeType: 'application/vnd.google-apps.spreadsheet' },
        blob, { fields: 'id' }
      );
      return r3.id;
    }
  } catch (e) { Logger.log('Drive v3 failed: ' + e.message); }

  // Method 2: Drive API v2
  try {
    if (typeof Drive !== 'undefined' && Drive.Files &&
        typeof Drive.Files.insert === 'function') {
      var r2 = Drive.Files.insert(
        { title: tempName, mimeType: 'application/vnd.google-apps.spreadsheet' },
        blob, { convert: true }
      );
      return r2.id;
    }
  } catch (e) { Logger.log('Drive v2 failed: ' + e.message); }

  // Method 3: UrlFetchApp multipart upload (no Advanced Service needed)
  return convertViaUrlFetch_(blob, tempName);
}

function convertViaUrlFetch_(blob, name) {
  var token    = ScriptApp.getOAuthToken();
  var boundary = 'shift_boundary_' + new Date().getTime();
  var metadata = JSON.stringify({
    name    : name,
    mimeType: 'application/vnd.google-apps.spreadsheet'
  });

  var body =
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    metadata + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n' +
    'Content-Transfer-Encoding: base64\r\n\r\n' +
    Utilities.base64Encode(blob.getBytes()) + '\r\n' +
    '--' + boundary + '--';

  var response = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id',
    {
      method             : 'POST',
      contentType        : 'multipart/related; boundary="' + boundary + '"',
      payload            : body,
      headers            : { Authorization: 'Bearer ' + token },
      muteHttpExceptions : true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error('HTTP ' + response.getResponseCode() + ': ' + response.getContentText());
  }
  var result = JSON.parse(response.getContentText());
  if (!result.id) throw new Error('No file ID returned: ' + response.getContentText());
  return result.id;
}


// ============================================================
//  HELPER: Extract shift rows from converted Google Sheet
// ============================================================
function extractShiftRows(sheetId, year, monthIndex) {
  var ss      = SpreadsheetApp.openById(sheetId);
  var ws      = ss.getSheets()[0];
  var data    = ws.getDataRange().getValues();
  if (data.length < 2) return [];

  var headerRow = data[0];

  // Find "Emp No" column
  var empNoIdx = -1;
  for (var i = 0; i < headerRow.length; i++) {
    if (normalise(headerRow[i]) === 'emp no') { empNoIdx = i; break; }
  }
  if (empNoIdx === -1) throw new Error('"Emp No" column not found.');

  // Find day columns (numeric headers 1–31)
  var dayCols = [];
  for (var j = 0; j < headerRow.length; j++) {
    if (j === empNoIdx) continue;
    if (DISCARD_HEADERS.indexOf(normalise(headerRow[j])) !== -1) continue;
    var raw = headerRow[j];
    var num = (typeof raw === 'number') ? raw : parseInt(raw, 10);
    if (!isNaN(num) && num >= 1 && num <= 31 && String(raw).trim() === String(num)) {
      dayCols.push({ colIdx: j, day: num });
    }
  }
  if (dayCols.length === 0) throw new Error('No numeric day columns (1–31) found.');

  var rows = [];
  for (var r = 1; r < data.length; r++) {
    var row   = data[r];
    var empNo = String(row[empNoIdx] || '').trim();
    if (!empNo) continue;

    for (var d = 0; d < dayCols.length; d++) {
      var dc    = dayCols[d];
      var shift = String(row[dc.colIdx] || '').trim();
      if (!shift) continue;

      // Validate date (skip e.g. April 31, Feb 30)
      var candidate = new Date(year, monthIndex, dc.day);
      if (candidate.getFullYear() !== year || candidate.getMonth() !== monthIndex) continue;

      var dateStr = Utilities.formatDate(candidate, 'Africa/Lagos', 'yyyy-MM-dd');
      rows.push([empNo, dateStr, shift]);
    }
  }
  return rows;
}


// ============================================================
//  HELPER: Normalise header string for comparison
// ============================================================
function normalise(val) {
  return String(val || '').trim().toLowerCase().replace(/\s+/g, ' ');
}


// ============================================================
//  HELPER: Month name → 0-based index
// ============================================================
function parseMonthName(name) {
  var MAP = {
    january:0, february:1, march:2,    april:3,
    may:4,     june:5,     july:6,     august:7,
    september:8, october:9, november:10, december:11
  };
  return MAP[name.toLowerCase()] !== undefined ? MAP[name.toLowerCase()] : -1;
}


// ============================================================
//  TRIGGER SETUP  — run ONCE manually
// ============================================================
/**
 * Installs a daily trigger at 1 AM WAT.
 * Requires the project timezone to be set to Africa/Lagos first
 * (Project Settings ⚙ → Time zone → Africa/Lagos).
 */
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {

    if (
    t.getHandlerFunction() === 'processShiftFiles' &&
    t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK
   ) {
    ScriptApp.deleteTrigger(t);
    }

  });

  ScriptApp.newTrigger('processShiftFiles')
    .timeBased()
    .atHour(1)
    .nearMinute(0)
    .everyDays(1)
    .create();

  Logger.log('✅ Daily trigger set: processShiftFiles() at ~1:00 AM Africa/Lagos.');
  }

