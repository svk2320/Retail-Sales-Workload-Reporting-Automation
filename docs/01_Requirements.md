# Requirements Specification

Retail Sales & Workload Reporting Automation

## 1. Input Data (from salesworkload_ORIGINAL.xlsx)

### Sheet: sales_figures (transactions, ~7,660 rows)

| Column        | Type             | Notes                            |
| ------------- | ---------------- | -------------------------------- |
| MonthYear     | text ("10.2016") | needs standardizing to real date |
| Time index    | int              | sequential month index           |
| Country       | text             | e.g. United Kingdom              |
| StoreID       | int              | join key to opening_schemes.id   |
| City          | text             | e.g. "London (I)"                |
| Dept_ID       | int              | 1–N                              |
| Dept. Name    | text             | e.g. Dry, Frozen, Fish, other    |
| HoursOwn      | decimal          | staff hours (owned)              |
| HoursLease    | decimal          | staff hours (leased)             |
| Sales units   | int              |                                  |
| Turnover      | int/decimal      | revenue                          |
| Customer      | often blank      | needs flag                       |
| Area (m2)     | decimal          |                                  |
| Opening hours | text             | "Type A" etc                     |

### Sheet: opening_schemes (store master, ~52 rows)

| Column                                  | Notes                                          |
| --------------------------------------- | ---------------------------------------------- |
| id                                      | join key = StoreID                             |
| Store name                              |                                                |
| Region                                  | actually holds Country in sample rows — verify |
| Scheme                                  | Type A/B/etc                                   |
| Month-by-month cols (10.2016...09.2017) | monthly hours-type metric                      |
| Cumulated cols                          | running total, same months                     |

This is our **transactions + store-master** pair — same structural role as
Sales/Product Master in the original spec, but real.

## 2. Target Workbook Structure (sheet-by-sheet)

1. **Raw_Data** — sales_figures pasted as-is, deliberately split into 3–4
   "weekly/monthly export" batches with intentional messiness introduced
   (Day/Step 2 task)
2. **Store_Master** — cleaned-up opening_schemes (id, Store name, Region,
   Scheme only — drop the monthly grid, not needed for lookups)
3. **Clean_Data** — VBA-cleaned Raw_Data + lookup columns:
   - XLOOKUP: Store name + Scheme from Store_Master by StoreID
   - VLOOKUP: Region from Store_Master by StoreID (separate column, proves
     VLOOKUP explicitly)
   - INDEX-MATCH: two-way — Store's opening-hours value for that row's
     MonthYear, pulled from the opening_schemes month-by-month grid
     (StoreID x MonthYear lookup)
4. **Data_Entry_Form** — manual "add a sale record" form with Data
   Validation dropdowns (Country, Dept. Name) and quantity/turnover > 0 rule
5. **Reports_Pivot** — PivotTables: Turnover by Country, Turnover by Dept,
   Month-over-month trend by StoreID. SUMIFS KPI cards above (e.g. "Total
   UK Turnover This Month") built as standalone formulas, not pivot-derived
6. **Alerts** — conditional formatting: low Turnover, blank/zero Customer
   flag, HoursOwn+HoursLease vs Area outlier check
7. **Cleaning_Log** — VBA-generated: rows removed, blanks flagged, run
   timestamp, before/after row counts
8. **Summary_Export** — one-click formatted report: KPIs + pivot snapshot,
   exportable to PDF

## 3. Macros to Build

| Macro             | Purpose                                                                    |
| ----------------- | -------------------------------------------------------------------------- |
| `CleanData`       | dedupe, fix MonthYear format, flag/remove invalid rows, write Cleaning_Log |
| `RefreshLookups`  | reapply XLOOKUP/VLOOKUP/INDEX-MATCH after cleaning changes row counts      |
| `RefreshPivots`   | refresh all PivotCaches                                                    |
| `ApplyFormatting` | reapply conditional formatting rules                                       |
| `GenerateSummary` | rebuild Summary_Export sheet                                               |
| `ExportToPDF`     | export Summary_Export dated PDF                                            |
| `RunFullReport`   | master macro — calls all above in sequence                                 |

## 4. Skills-to-Feature Map (for resume/README table)

| Skill                  | Where it lives                                                                                         |
| ---------------------- | ------------------------------------------------------------------------------------------------------ |
| PivotTables            | Reports_Pivot                                                                                          |
| XLOOKUP                | Clean_Data (Store name, Scheme)                                                                        |
| VLOOKUP                | Clean_Data (Region)                                                                                    |
| INDEX-MATCH            | Clean_Data (monthly opening-hours two-way lookup)                                                      |
| SUMIFS                 | Reports_Pivot KPI cards                                                                                |
| Conditional Formatting | Alerts                                                                                                 |
| Data Validation        | Data_Entry_Form                                                                                        |
| VBA/Macros             | CleanData, RefreshLookups, RefreshPivots, ApplyFormatting, GenerateSummary, ExportToPDF, RunFullReport |

## 5. Before/After Proof Requirement

Time a full manual pass (clean one messy batch, add lookups, refresh
pivots, apply formatting, export) vs. one click of `RunFullReport`.
Record both. This produces the real X→Y numbers for the resume bullet
(no invented numbers).

## 6. Next Step After This Doc

Introduce deliberate messiness into a working copy of Raw_Data (this is
the "before" state) — nothing else changes yet.
