# 6. Leave Calculation Method (Official)

Source file: `leave calculation.docx` (also under `/home/workdir/artifacts/leave_calculation.docx`)

## 6.1 Overview

Leave utilization is driven by **Calculation Method** on each leave type in `Sys_LeavePolicies`:

- **Shift Roaster** (`ShiftRoaster`)
- **Actual Days** (`ActualDays`)

## 6.2 Methods

### A. Shift Roaster

For every date from Start Date to End Date (**inclusive**), read the employee’s assigned shift and apply the **Shift Roaster Multiplier**. The sum is **Leave Utilized**.

### B. Actual Days

**Leave Utilized** = inclusive calendar days (`End − Start + 1`).  
Shift codes are ignored; **every day counts as 1**, including `O`.

## Official multiplier table

| Shift | Shift Roaster Multiplier | Actual Days Multiplier |
|-------|--------------------------|------------------------|
| G | 1 | 1 |
| A | 1.5 | 1 |
| B | 1.5 | 1 |
| D | 1.5 | 1 |
| N | 1.5 | 1 |
| O | 0 | 1 |

## Example — FRT7524, 01-Jan-2026 to 05-Jan-2026

| Date | Shift | Annual Leave (ShiftRoaster) | Leave of Absence (ActualDays) |
|------|-------|----------------------------|-------------------------------|
| 01-Jan-2026 | D | 1.5 | 1 |
| 02-Jan-2026 | D | 1.5 | 1 |
| 03-Jan-2026 | O | 0 | 1 |
| 04-Jan-2026 | O | 0 | 1 |
| 05-Jan-2026 | G | 1 | 1 |
| **Total** | | **4** | **5** |

## 6.6 Leave Balance Deduction

After computing utilization:

1. Deduct from the employee’s available balance for that leave type (and from **Annual** when `Deduct from` = Annual Leave).
2. Reflect utilization in balances, reports, and leave history.
3. Previous-year balance expires per **Carry Forward Deadline** on the policy.

## Engine behaviour in this repo

| Situation | Behaviour |
|-----------|-----------|
| Method = ShiftRoaster | Day loop; multipliers as in table above (G=1, A/B/D/N=1.5, O=0). Code `M` treated like A/B/D/N → 1.5 |
| Method = ActualDays | Inclusive calendar days only |
| No shift row for a date | Default **Mon–Fri = G**, **Sat–Sun = O**, then apply Shift Roaster multipliers |
| Casual / Examination | Utilization calculated per their method; balance deducts from **Annual** as well |
