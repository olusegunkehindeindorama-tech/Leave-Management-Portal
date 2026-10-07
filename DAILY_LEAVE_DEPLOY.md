# Daily leave model — deploy steps

1. **Backup** the Google Spreadsheet.
2. Add scripts in this **load order**:
   - DateUtils.gs
   - Entitlement and Utilization.gs (policies / master load)
   - **DailyLeaveModel.gs**
   - **DailyEntitlement.gs** (overrides `calculateLeaveUtilized`)
   - **DailyCleanup.gs**
   - **DailyImportBridge.gs**
   - Shift.gs, imports, Code.gs, Index.html
3. Run once:
   ```javascript
   migrateTblLeaveToDaily()
   ```
4. Run:
   ```javascript
   recalculateAllLeaveUtilized()
   ```
5. Disable/remove obsolete base64 files: `IndexB64P*`, `IndexHtmlP*`, `IndexHtmlLoader.gs`
6. After migration, do **not** run old range-based Leave Cleanup overlap logic.

## Expected result for a 3–7 Apr leave
Five rows, same Entry Code, Leave Date = each day, Leave Utilized from shift multipliers, Entitlement Year from carry-forward rules.

## Dedup key
`Emp ID | Leave Date | Leave Type` — one day of a leave type per employee.
