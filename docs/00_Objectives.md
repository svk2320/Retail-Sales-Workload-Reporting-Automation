# Project: Retail Sales & Workload Reporting Automation

(Excel + VBA)

## 1. Business Problem (simulated)

A multi-country retail chain (UK stores, multiple departments per store)
receives raw monthly sales exports. Someone manually:

- cleans the export (blank fields, inconsistent formats)
- cross-references store opening-hours/scheme data
- builds pivot reports (sales by store, by department, by month)
- flags underperforming stores/departments
- formats and sends a summary report

This is repeated every reporting cycle — pure manual grind.

## 2. What This Project Automates

End-to-end: raw export -> cleaned data -> lookups -> pivot reports ->
conditional-formatted alerts -> one-click formatted summary (screen/PDF).

## 3. Source Data (real, not simulated)

File: `salesworkload_ORIGINAL.xlsx`

- **Sheet 1: `sales_figures`** — transactional-style monthly records per
  Store/Department: MonthYear, Time index, Country, StoreID, City, Dept_ID,
  Dept. Name, HoursOwn, HoursLease, Sales units, Turnover, Customer,
  Area (m2), Opening hours. ~7,660 rows.
- **Sheet 2: `opening_schemes`** — one row per store, with Region, Scheme
  type, and month-by-month + cumulated opening-hours-type metric across
  Oct 2016–Sep 2017. ~52 rows.

This naturally gives us a "transactions table" + a "store master table" —
which is exactly the structure needed for lookups (XLOOKUP/VLOOKUP/
INDEX-MATCH), same as the original spec, but grounded in real data instead
of an invented Product Master.

## 4. Skills to Demonstrate (resume-mapped)

- PivotTables (sales by store, by dept, by month trend)
- XLOOKUP (pull Region/Scheme from opening_schemes into sales data by StoreID)
- VLOOKUP (legacy-compatible lookup, separate use case)
- INDEX-MATCH (two-way lookup, e.g. store's opening-hours value for a
  specific month)
- SUMIFS (conditional totals — e.g. turnover by Country AND month)
- Conditional Formatting (flag low turnover, negative/low customer count,
  outlier sales)
- Data Validation (dropdowns for Country/Dept on a manual entry form)
- VBA/Macros (cleaning, lookup refresh, pivot refresh, report generation,
  one-click "Refresh All")

## 5. Scope Boundaries (keep it tight)

- One dataset (this file), no invented data beyond deliberately-introduced
  messiness for the "before" story
- One recurring report cycle simulated (monthly), not a live/streaming system
- No external database or ERP integration

## 6. Deliverables

- `.xlsm` macro-enabled workbook, commented VBA
- README (business problem, architecture, skills table, before/after timing)
- Screenshots: cleaning log, lookup columns, pivot dashboard, conditional
  formatting, data validation, exported PDF
- Before/after timing proof (manual vs. one-click)

## 7. Target Resume Bullet

> "Built a VBA-automated Excel reporting system integrating XLOOKUP,
> VLOOKUP, INDEX-MATCH, SUMIFS, PivotTables, and Data Validation across a
> multi-sheet retail sales workbook — reduced manual reporting from ~1
> hour to under 12 seconds."

## 8. Status

- [x] Objectives defined
- [x] Requirements spec
- [x] Raw data messiness introduced (before-state)
- [x] Cleaning macro
- [x] Lookup layer
- [x] Pivot layer
- [x] Conditional formatting
- [x] Data validation form
- [x] One-click master macro
- [x] README + screenshots + timing proof

Project complete — see root [`README.md`](../README.md) for final results.
