/**
 * ImportDateBridge.gs — load AFTER DB Import.gs and Excel Leave Import.gs
 * Ensures Start/End use DateUtils (2h+ nudge). Safe if DateUtils.gs is present.
 *
 * Prefer updating DB Import.gs / Excel Leave Import.gs fully from project artifacts:
 *   DB_Import.gs and Excel_Leave_Import.gs in the project folder.
 */

(function () {
  if (typeof toCalendarDate_ === 'function') {
    parseLeaveDate_ = function (val) {
      return toCalendarDate_(val);
    };
  }
  if (typeof normalizeImportDate_ !== 'function') {
    normalizeImportDate_ = function (val) {
      if (typeof toCalendarDate_ === 'function') return toCalendarDate_(val);
      return parseLeaveDate_(val);
    };
  }
  if (typeof sheetDateForWrite_ !== 'function') {
    sheetDateForWrite_ = function (val) {
      if (typeof toSheetDateValue_ === 'function') {
        var v = toSheetDateValue_(val);
        return v === '' ? null : v;
      }
      var d = normalizeImportDate_(val);
      if (!d) return null;
      return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
    };
  }
})();
