/**
 * ============================================================
 *  SHIFT DATA PROCESSOR  — Google Apps Script
 * ============================================================
 *  - Reads "YYYY Month.xlsx" files from SOURCE_FOLDER_ID
 *  - Creates temporary Google Sheets INSIDE the same folder
 *  - After successful extract: trashes the temp GS and the original .xlsx
 *  - INCREMENTAL processing + daily trigger support
 * ============================================================
 */

var SOURCE_FOLDER_ID = '1WX2RK3b5lPegwNrLzVrWs7vJ-_sWQnF1';
var DEST_FILE_ID     = '13ahfJJLgh3lBO4VPSWG6nRau_9MIPrryfNrxPs4veKA';
var DEST_SHEET_NAME  = 'tblShift';
var QUEUE_KEY = 'shiftProcessingQueue';
var WORKER_TRIGGER = 'processShiftFiles';
var DISCARD_HEADERS = ['name', 'work group', 'shift/ general', 'shift/general',
                       'function dept', 'dept'];
var PROP_KEY = 'shiftFilesMeta';

function processShiftFiles() {
  var props = PropertiesService.getScriptProperties();
  var queue = JSON.parse(props.getProperty(QUEUE_KEY) || '[]');

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

  if (queue.length === 0) {
    Logger.log('No new or modified files.');
    deleteWorkerTriggers_();
    return { success: true, message: 'No new or modified shift files.' };
  }

  var entry = queue.shift();
  props.setProperty(QUEUE_KEY, JSON.stringify(queue));
  Logger.log('Processing: ' + entry.fileName);
  processSingleFile_(entry);

  if (queue.length > 0) {
    scheduleWorker_();
    Logger.log('Remaining files: ' + queue.length);
  } else {
    deleteWorkerTriggers_();
    Logger.log('Queue completed.');
  }

  return { success: true, message: 'Processed ' + entry.fileName + '. Remaining in queue: ' + queue.length };
}

function processSingleFile_(entry) {
  var props = PropertiesService.getScriptProperties();
  var storedMeta = JSON.parse(props.getProperty(PROP_KEY) || '{}');
  var destSS = SpreadsheetApp.openById(DEST_FILE_ID);
  var destSheet = destSS.getSheetByName(DEST_SHEET_NAME);
  if (!destSheet) destSheet = destSS.insertSheet(DEST_SHEET_NAME);

  var tempId = null;
  var originalFile = null;

  try {
    if (entry.isModified) {
      Logger.log('Removing existing rows for month...');
      removeMonthRows_(destSheet, entry.year, entry.monthIndex);
    }

    originalFile = DriveApp.getFileById(entry.fileId);
    tempId = convertXlsxToGSheetInFolder_(originalFile);

    var rows = extractShiftRows(tempId, entry.year, entry.monthIndex);
    writeChunked_(destSheet, rows);

    storedMeta[entry.fileId] = {
      modified: entry.modified,
      year: entry.year,
      month: entry.monthIndex
    };
    props.setProperty(PROP_KEY, JSON.stringify(storedMeta));

    Logger.log('Finished: ' + entry.fileName + ' (' + rows.length + ' rows)');

    // Delete original xlsx after successful conversion + extract
    try {
      originalFile.setTrashed(true);
      Logger.log('Original xlsx trashed: ' + entry.fileName);
    } catch (delErr) {
      Logger.log('Could not trash original xlsx: ' + delErr.message);
    }

  } finally {
    if (tempId) {
      try { DriveApp.getFileById(tempId).setTrashed(true); } catch (e) {}
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
        row[1] = Utilities.formatDate(dateVal, 'Africa/Lagos', 'yyyy-MM-dd');
      }
      kept.push(row);
    }
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['Emp ID', 'Date', 'Shift']]);
  writeChunked_(sheet, kept);
}

function scheduleWorker_() {
  ScriptApp.newTrigger(WORKER_TRIGGER).timeBased().after(1000).create();
}

function deleteWorkerTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === WORKER_TRIGGER) ScriptApp.deleteTrigger(t);
  });
}

function clearStateAndReprocess() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_KEY);
  props.deleteProperty(QUEUE_KEY);
  deleteWorkerTriggers_();

  var destSS = SpreadsheetApp.openById(DEST_FILE_ID);
  var destSheet = destSS.getSheetByName(DEST_SHEET_NAME);
  if (destSheet) {
    destSheet.clearContents();
    destSheet.getRange(1, 1, 1, 3).setValues([['Emp ID', 'Date', 'Shift']]);
  }
  processShiftFiles();
}

function writeChunked_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  var CHUNK_SIZE = 10000;
  var startRow = sheet.getLastRow() + 1;
  for (var i = 0; i < rows.length; i += CHUNK_SIZE) {
    var chunk = rows.slice(i, i + CHUNK_SIZE);
    sheet.getRange(startRow, 1, chunk.length, 3).setValues(chunk);
    startRow += chunk.length;
  }
}

/**
 * Convert .xlsx → temporary Google Sheet placed INSIDE SOURCE_FOLDER_ID.
 */
function convertXlsxToGSheetInFolder_(file) {
  var blob = file.getBlob();
  var tempName = '_TEMP_SHIFT_' + file.getId() + '_' + new Date().getTime();
  var folderId = SOURCE_FOLDER_ID;

  // Method 1: Drive API v3 with parents
  try {
    if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.create === 'function') {
      var r3 = Drive.Files.create(
        {
          name: tempName,
          mimeType: 'application/vnd.google-apps.spreadsheet',
          parents: [folderId]
        },
        blob,
        { fields: 'id' }
      );
      return r3.id;
    }
  } catch (e) { Logger.log('Drive v3 failed: ' + e.message); }

  // Method 2: Drive API v2
  try {
    if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.insert === 'function') {
      var r2 = Drive.Files.insert(
        {
          title: tempName,
          mimeType: 'application/vnd.google-apps.spreadsheet',
          parents: [{ id: folderId }]
        },
        blob,
        { convert: true }
      );
      return r2.id;
    }
  } catch (e) { Logger.log('Drive v2 failed: ' + e.message); }

  // Method 3: UrlFetch with parents
  return convertViaUrlFetchInFolder_(blob, tempName, folderId);
}

function convertViaUrlFetchInFolder_(blob, name, folderId) {
  var token = ScriptApp.getOAuthToken();
  var boundary = 'shift_boundary_' + new Date().getTime();
  var metadata = JSON.stringify({
    name: name,
    mimeType: 'application/vnd.google-apps.spreadsheet',
    parents: [folderId]
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
      method: 'POST',
      contentType: 'multipart/related; boundary="' + boundary + '"',
      payload: body,
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error('HTTP ' + response.getResponseCode() + ': ' + response.getContentText());
  }
  var result = JSON.parse(response.getContentText());
  if (!result.id) throw new Error('No file ID returned: ' + response.getContentText());
  return result.id;
}

function extractShiftRows(sheetId, year, monthIndex) {
  var ss = SpreadsheetApp.openById(sheetId);
  var ws = ss.getSheets()[0];
  var data = ws.getDataRange().getValues();
  if (data.length < 2) return [];

  var headerRow = data[0];
  var empNoIdx = -1;
  for (var i = 0; i < headerRow.length; i++) {
    if (normalise(headerRow[i]) === 'emp no') { empNoIdx = i; break; }
  }
  if (empNoIdx === -1) throw new Error('"Emp No" column not found.');

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
    var row = data[r];
    var empNo = String(row[empNoIdx] || '').trim();
    if (!empNo) continue;

    for (var d = 0; d < dayCols.length; d++) {
      var dc = dayCols[d];
      var shift = String(row[dc.colIdx] || '').trim();
      if (!shift) continue;

      var candidate = new Date(year, monthIndex, dc.day);
      if (candidate.getFullYear() !== year || candidate.getMonth() !== monthIndex) continue;

      var dateStr = Utilities.formatDate(candidate, 'Africa/Lagos', 'yyyy-MM-dd');
      rows.push([empNo, dateStr, shift]);
    }
  }
  return rows;
}

function normalise(val) {
  return String(val || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function parseMonthName(name) {
  var MAP = {
    january: 0, february: 1, march: 2, april: 3,
    may: 4, june: 5, july: 6, august: 7,
    september: 8, october: 9, november: 10, december: 11
  };
  return MAP[name.toLowerCase()] !== undefined ? MAP[name.toLowerCase()] : -1;
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processShiftFiles' &&
        t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processShiftFiles')
    .timeBased()
    .atHour(1)
    .nearMinute(0)
    .everyDays(1)
    .create();
  Logger.log('Daily trigger set: processShiftFiles() at ~1:00 AM Africa/Lagos.');
}
