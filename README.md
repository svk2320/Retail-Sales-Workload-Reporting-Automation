# Retail Sales & Workload Reporting Automation

Excel + VBA project — automated data cleaning, lookups, PivotTable
reporting, conditional formatting, and one-click PDF export.

## Business Problem

A multi-country retail chain (11 countries, 50 stores, department-level
sales tracking) receives a new sales export every reporting cycle.
Historically, someone would manually:

- clean the raw export (duplicates, inconsistent formats, blank fields)
- cross-reference store master data (region, scheme type)
- build PivotTable reports (by country, department, month)
- flag underperforming stores/departments
- format and export a summary report

This project automates that entire pipeline into a single macro.

## Real Data, Simulated "Before" State

The underlying dataset (`salesworkload_ORIGINAL.xlsx`) is real retail
sales/workload data — not invented. To build a genuine before/after
automation story, the data was split into 4 simulated monthly export
batches (Jan–Apr 2017), with 2 of the 4 deliberately corrupted
(duplicate rows, inconsistent date formats, blank required fields,
inconsistent text casing) to represent a realistic "messy raw export."
See `02_Why_Messy_Batches.md` for full reasoning.

## Architecture

| Sheet             | Purpose                                                                            |
| ----------------- | ---------------------------------------------------------------------------------- |
| `Raw_Data`        | Combined raw import of all 4 batches (messy, "before" state)                       |
| `Store_Master`    | Clean 50-store lookup table (id, Store name, Region, Scheme)                       |
| `Clean_Data`      | Output of `CleanData` — deduplicated, standardized, flagged, with 3 lookup columns |
| `Data_Entry_Form` | Manual entry row with dropdown + positive-number validation                        |
| `Reports_Pivot`   | 3 manually-built PivotTables + 3 SUMIFS KPI cards                                  |
| `Alerts`          | Live mirror of `Clean_Data` with 3 conditional formatting rules                    |
| `Cleaning_Log`    | Audit trail: rows in/out, duplicates removed, rows flagged, run timestamp          |
| `Summary_Export`  | One-page dashboard: KPIs + all 3 pivot snapshots, exported to PDF                  |

## Skills Demonstrated

| Skill                  | Where it lives                                                                               |
| ---------------------- | -------------------------------------------------------------------------------------------- |
| PivotTables            | `Reports_Pivot` — built manually (Country, Department, Month trend)                          |
| XLOOKUP                | `Clean_Data` col P — with automatic fallback to INDEX-MATCH if unsupported (Excel <365/2021) |
| VLOOKUP                | `Clean_Data` col Q — Region lookup                                                           |
| INDEX-MATCH            | `Clean_Data` col R — Scheme lookup                                                           |
| SUMIFS                 | `Reports_Pivot` KPI cards — single- and multi-condition totals                               |
| Conditional Formatting | `Alerts` sheet — 3 rules, applied via VBA                                                    |
| Data Validation        | `Data_Entry_Form` — dropdowns + numeric range restriction                                    |
| VBA/Macros             | 7 macros (see below), chained into one master macro                                          |

## VBA Macros

Full source in `VBA_Macros.bas`.

| Macro             | Purpose                                                                                                              |
| ----------------- | -------------------------------------------------------------------------------------------------------------------- |
| `CleanData`       | Dedupes, standardizes dates, normalizes text casing, excludes rollup rows, flags invalid rows, writes `Cleaning_Log` |
| `RefreshLookups`  | Writes/refreshes the 3 lookup columns (auto-detects XLOOKUP support)                                                 |
| `ApplyFormatting` | Mirrors `Clean_Data` into `Alerts`, applies conditional formatting                                                   |
| `RefreshPivots`   | Refreshes all PivotCaches on `Reports_Pivot`                                                                         |
| `GenerateSummary` | Builds the `Summary_Export` dashboard                                                                                |
| `ExportToPDF`     | Exports `Summary_Export` as a dated PDF                                                                              |
| `RunFullReport`   | Master macro — chains all of the above in sequence, one click                                                        |

## Real Bugs Found & Fixed (worth knowing for an interview)

This project surfaced 3 genuine bugs during development, not scripted —
full detail in `04_Note_MonthYear_Text_Format.md` and inline comments in
`VBA_Macros.bas`:

1. **Excel auto-converting text dates.** Values like `"Feb-17"` were
   silently converted by Excel into a real Date/serial number on write,
   corrupting the intended text format.
2. **A rollup row hiding in the "real" data.** The dataset contained a
   `Dept. Name = "all"` value — a store-level grand-total row, not a
   real department — silently inflating department/month totals by
   ~7.1B before it was caught and excluded in `CleanData`.
3. **VBA's `IsDate()` misreading ambiguous text.** `IsDate("Apr-17")`
   returns `True`, but VBA interprets it as "April 17th of the current
   system year" (reading `17` as a day, not a year) — producing
   `04.2026` instead of `04.2017`. Fixed by checking explicit text
   patterns before any date-type fallback, and using
   `VarType(val) = vbDate` instead of `IsDate()`.

## Before/After

- **Automated (`RunFullReport`):** 11.8 seconds (verified, real timing
  from the macro's own popup)
- **Manual baseline:** ~1 hour (reasoned estimate — this is a recurring
  task, redone every time a new month's data arrives: cleaning, 3
  lookup types across ~3,000 rows, 3 pivots, conditional formatting, the
  entry form, and the summary/PDF export, all by hand)

**Resume bullet:**

> Built a VBA-automated Excel reporting system integrating
> XLOOKUP/INDEX-MATCH, VLOOKUP, SUMIFS, PivotTables, and Data Validation
> across a multi-sheet retail sales workbook — reduced an estimated ~1
> hour of recurring manual reporting to under 12 seconds with a
> one-click macro.

## Repository Structure

```
├── README.md
├── LICENSE
├── RetailSalesWorkloadAutomation.xlsm   ← the working macro-enabled workbook
├── VBA_Macros.bas                       ← all 7 macros, standalone source file
├── Sales_Summary_2026-09-01.pdf         ← sample exported report
│
├── docs/
│   ├── 00_Objectives.md                 ← business problem, scope, target resume bullet
│   ├── 01_Requirements.md               ← functional spec (sheets, macros, skills map)
│   ├── 02_Why_Messy_Batches.md          ← why/how the "before" state was simulated
│   ├── 03_Build_Workbook_Shell_Notes.md ← what the workbook-assembly script did
│   ├── 04_Note_MonthYear_Text_Format.md ← the date-handling bug story in detail
│   └── Dataset Source.txt               ← source data attribution
│
├── project progress status/
│   └── PROJECT_STATUS (0–5).md          ← full build log, in order, with every decision and fix
│
├── screenshots/                         ← Cleaning Log, Clean Data, Data Entry Form,
│                                            Raw Data, Reports Pivot, Store Master, Summary Export
│
├── scripts/
│   ├── make_messy_batches.py            ← one-time script that generated the messy batch files
│   └── build_workbook_shell.py          ← one-time script that assembled the initial workbook shell
│
├── source_data/
│   ├── salesworkload_ORIGINAL.xlsx      ← original real source data (untouched)
│   └── export_batch_*.xlsx              ← the 4 simulated monthly export batches
│
└── archive/                             ← earlier workbook/macro versions, kept for history
```

> Note: `archive/` contains earlier drafts of the workbook and macros, kept for build-history reference — not the current deliverable. The current files are in the repository root.

## Screenshots

See `/screenshots` : Cleaning_Log after a run, lookup
columns in `Clean_Data`, the 3 PivotTables, conditional formatting on
`Alerts`, the Data Validation dropdown in action, and the exported PDF.

## License

© 2026 Vasanth Kumar S. All Rights Reserved.

This project is publicly available for portfolio and evaluation purposes.
Please do not copy, redistribute, modify, or reuse the source code without
permission.
