# Note: Why MonthYear Stays as Text (Not Converted to a Date)

## The bug we found
VBA's `IsDate()` misinterpreted text like `"Apr-17"` as "April 17th,
current year" instead of "April 2017" — a genuine locale/date-parsing
ambiguity, not a typo. This silently corrupted values (e.g. produced
`04.2026` instead of `04.2017`).

## The fix
Check explicit text patterns (`MM.YYYY`, `YYYY-MM`, `Mon-YY`) before
falling back to any date-type detection. Use `VarType(val) = vbDate`
instead of `IsDate()` to avoid false positives — `VarType` only catches
values that are *actually* a Date type, not text VBA merely guesses
looks date-like.

The cleaned `MonthYear` column is then written with
`NumberFormat = "@"` (force Text) before the value is set, so Excel
can't silently re-convert it into a date again later.

## Why "Number Stored as Text" is expected, not a problem
Excel flags the cell with a warning triangle ("Number Stored as Text").
This is intentional and correct — do NOT convert it to a real date type.

1. `MonthYear` values like `04.2017` represent a **month-level period
   label**, not a specific calendar date — there's no real "day"
   component, it's a grouping key (like "Q1" or "Week 3").
2. Converting it to a real Excel date would force Excel to invent an
   arbitrary day (usually the 1st) — reintroducing the exact ambiguity
   that caused the original bug.
3. PivotTables already group and sort it correctly as text
   (`01.2017` → `02.2017` → `03.2017` → `04.2017`) because it's
   zero-padded, so alphabetical sort = chronological sort.
4. The warning is a suggestion, not an error. Safe to ignore, or select
   the column and choose "Ignore Error" to silence it (cosmetic only).

## When a real date type WOULD be needed
Only if day-level precision were required, or date-specific functions
(`EDATE`, `WORKDAY`, days-between calculations) were needed. Not the
case here — only month-level grouping/filtering is ever used.

**Conclusion: leave `MonthYear` as text. This is the fix, not a side
effect to clean up further.**
