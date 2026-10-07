# Daily leave model — what to keep / delete

## KEEP in Apps Script project
| File | Role |
|------|------|
| DateUtils.gs | Noon / WAT date helpers |
| DailyLeaveModel.gs | Schema, migrate, expand range→days |
| DailyImportBridge.gs | importRangeLeavesAsDaily_, UI submit helper |
| **DailyCleanup.gs** | **Only** cleanup pipeline + schedule |
| **zz_DailyRecalc.gs** | **Only** recalculate util + entitlement (loads last) |
| Entitlement and Utilization.gs | Policies, multipliers, master loaders |
| Shift.gs, imports, Code.gs, UiApi.gs, Index.html, Cache.gs, EmployeeSync.gs, Leave balance.gs | App |

## DELETE from Apps Script project (must remove)
| File | Why |
|------|-----|
| **Leave Cleanup.gs** | Old range overlap/dedupe — this is what was still running your “recalc” |
| LeaveDateWriteFix.gs | Superseded |
| CarryForwardFix.gs | Superseded by zz_DailyRecalc |
| ImportDateBridge.gs | Range Start/End bridge |
| IndexB64P*, IndexHtmlP*, IndexHtmlLoader.gs | Base64 UI |
| DiagnoseTblLeave.gs | Range diagnostic |

## Run only these
```javascript
migrateTblLeaveToDaily()       // once
recalculateAllLeaveUtilized()  // must log: zz_DailyRecalc or DAILY RECALC START
runLeaveCleanupPipeline()      // daily dedupe + recalc
```

## Dedup key
`Emp ID | Leave Date` only — one leave row per employee per calendar day.
