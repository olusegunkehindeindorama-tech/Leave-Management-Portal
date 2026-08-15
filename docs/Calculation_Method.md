# Leave Calculation Method (from Sys_LeavePolicies + live rules)

## Calculation Method column
- **ShiftRoaster**: Count each calendar day using `tblShift` codes. If no roster row for that day, default **Mon–Fri = G (1 day)**, **Sat–Sun = O (0)**.
- **ActualDays**: Inclusive calendar days (end − start + 1). No roster.

## Multiplier column (Yes/No on Sys_LeavePolicies)
When **Multiplier = Yes** (typically Annual Leave):
- Shift codes **A, B, D, N, M** → **1.5** days each
- **G** → **1**
- **O** (and other off) → **0**
When **Multiplier = No**: working codes count as **1**, O as **0**.

## Deduct from
- **Casual Leave** → deducts from **Annual Leave** balance (also capped by Casual entitlement)
- **Examination Leave** → deducts from **Annual Leave** balance
- Comment on balance for Annual: balance after sum(Annual + Casual + Examination usage)

## Carry forward
- From StartingBal `{prevYear} Balance` for Annual only
- Expires on **Carry Forward Deadline** (e.g. 30 Jun current year)
- Usage consumes carry-forward first, then current year entitlement

## Defaults when no roster
Mon–Fri treated as **G**, Saturday/Sunday as **O**.
