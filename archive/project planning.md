
## Project: Sales & Inventory Reporting Automation (Excel + VBA)

**The business problem to simulate:** A small business (pick retail, or reuse a supermarket/FMCG dataset from Kaggle) gets a raw sales export every day/week and someone manually builds the same report over and over — cleaning it, building pivot tables, formatting, emailing it out. You're automating that grind.

**What to build:**

1. **Raw data intake** — a folder where new CSV/Excel sales exports get dropped (simulate a few weeks of files)
2. **A macro-driven cleaning step** — VBA that strips duplicates, fixes date formats, flags blank/invalid rows, standardizes product/category names
3. **Auto-refreshing PivotTables** — Sales by region, top products, month-over-month trend — refreshed via a macro button, not manually rebuilt each time
4. **Conditional formatting rules** — auto-highlight low stock, negative margins, or outlier sales spikes
5. **One-click summary report** — a macro that generates a formatted "Weekly Sales Summary" sheet (or exports to PDF) with the pivots + a few KPIs, ready to email
6. **A "Refresh All" button** — single macro that re-runs cleaning → pivot refresh → report generation, so the whole thing goes from raw file to finished report in one click

**The story for your resume (this is the actual point):**
> "Built a VBA-automated Excel reporting tool that reduced manual weekly sales report preparation from ~2 hours to under 5 minutes, including automated data cleaning, PivotTable refresh, and formatted summary export."

That "before/after time saved" framing is exactly what MIS/Reporting Analyst interviewers want to hear — it's not about the cleverness of the macro, it's about the efficiency story.

**Keep scope tight:**
- Don't try to build a full ERP-like system
- One dataset, one recurring report, fully automated end-to-end — that's more convincing than five half-built features
- Record a short before/after (screen recording or screenshots of manual process vs. one-click) — great for a portfolio page or interview demo

Want me to help you find a suitable public dataset for this, or start scaffolding the VBA macro logic?

Good options here. My pick, and why:

**Recommended dataset: "Sales and workload in retail industry"**
`https://www.kaggle.com/datasets/dgluesen/sales-and-workload-data-from-retail-industry`
It's explicitly described as a real-world-recreated spreadsheet for data preprocessing — which fits perfectly, since messy real-world formatting is exactly what your VBA cleaning macro needs to demonstrate value against.

**Backup option: "Retail Sales Dataset" (mohammadtalib786)**
Simple, transactional, easy to split into weekly batches — good if you want more control over simulating "new export every week."

## Suggested build sequence

**Day 1 — Setup & fake the "weekly export" scenario**
- Download the dataset, split it into 3-4 chunks by date (simulating weekly exports)
- Deliberately mess a couple of them up (duplicate rows, inconsistent date formats, blank cells, mismatched category names) — this is what your "before" state needs to justify the automation

**Day 2 — Build the cleaning macro**
- VBA sub that: removes duplicates, standardizes dates, flags/removes blank critical fields, trims whitespace, normalizes text case
- Log what it cleaned (a "Cleaning Log" sheet — rows removed, fixes applied) — this doubles as your data-quality-report evidence, like your other two projects

**Day 3 — Pivot + formatting automation**
- Macro-driven PivotTable refresh: sales by region/category, top products, week-over-week trend
- Conditional formatting macro: flag low stock / margin outliers

**Day 4 — One-click report + before/after proof**
- "Refresh All" master macro chaining everything
- Export a formatted summary sheet (or PDF) — this is your polished deliverable
- Time yourself doing it manually first, then with the macro, so you have a real before/after number for your resume line

Want me to start scaffolding the actual VBA code for the cleaning macro first?

## Excel Sales Reporting Automation — Full Project Spec

### 1. Objective
Simulate a recurring business report that currently takes manual effort each week, and automate it end-to-end with VBA so it runs in one click.

### 2. Business scenario (write this down, it's your framing)
A retail business receives a raw sales export weekly. Someone manually cleans it, builds pivot tables, formats a summary, and sends it out — roughly 2 hours of repetitive work each time. Your project replaces that with a one-click macro.

### 3. Dataset
- Kaggle: "Sales and workload in retail industry" (dgluesen) — real-world-recreated, good for showing messy-data cleaning
- Backup: "Retail Sales Dataset" (mohammadtalib786) — simpler, transactional
- Split into 3-4 chunks to simulate separate weekly exports
- Deliberately introduce messiness in 1-2 chunks: duplicate rows, inconsistent date formats, blank required fields, inconsistent text casing/category names — you need a believable "before" state

### 4. Tools needed
- Microsoft Excel (with Developer tab enabled for VBA/macros)
- The split CSV/XLSX files
- Nothing else — no external software required

### 5. What to build (in order)

**A. Raw data intake structure**
- A folder simulating incoming weekly exports
- A "Raw Data" sheet where each week's file gets pasted/imported

**B. Data cleaning macro**
- Remove duplicate rows
- Standardize date formats
- Flag/remove blank or invalid rows (missing price, missing product ID, etc.)
- Trim whitespace, normalize text case (product/category names)
- Write a "Cleaning Log" sheet: rows removed, issues fixed, count before/after — this is your audit trail, same spirit as the data-quality reports in your other two projects

**C. Reporting layer**
- PivotTables: Sales by region, Top products, Month-over-month/week-over-week trend
- A macro to refresh all pivots in one call (no manual "refresh" clicking)

**D. Visual layer**
- Conditional formatting: flag low stock, negative margin, sales spikes/outliers
- Applied via macro so it reapplies automatically on new data, not manually set each time

**E. One-click master report**
- A "Refresh All" button/macro that chains: import → clean → refresh pivots → apply formatting → generate summary sheet
- Summary sheet should be presentation-ready — a few KPIs at top, then the pivot charts, formatted for either screen or PDF export
- Optional: macro to export summary as PDF automatically

**F. Before/after proof (this is what sells the project)**
- Time yourself doing the process manually once
- Time the one-click macro run
- Screenshot or short screen recording of both
- This becomes your resume line: "reduced report prep from ~X hours to Y minutes"

### 6. Deliverables to have at the end
- The working `.xlsm` file (macro-enabled) with all macros documented (comments in code)
- A short README (like your other two projects) explaining: business problem, what was automated, before/after time, tech used
- Screenshots of the cleaning log, the pivot dashboard, and the conditional-formatted view
- The before/after timing evidence

### 7. Resume bullet to end up with
> "Built a VBA-automated Excel reporting tool that reduced manual weekly sales report preparation from ~2 hours to under 5 minutes — automating data cleaning, PivotTable refresh, conditional formatting, and summary generation."

### 8. Time estimate
3-4 days, working solo, if you keep scope exactly as above and don't add extra features.

Want me to start writing the actual VBA macro code next, or the README structure first?

Understood — let's expand it into a genuine "master" Excel project that hits every skill on your resume, not just automation. Here's the full spec.

## Excel Master Project — Full Spec (Sales & Inventory Command Center)

### 1. Business scenario
A retail business gets weekly sales + inventory exports. Someone manually cleans data, cross-references product/pricing info across sheets, builds pivot reports, tracks stock issues, and sends a formatted summary — all by hand. You're building the one-workbook system that does all of it.

### 2. Skills to demonstrate (map every resume bullet to a feature)

| Resume skill | Where it shows up |
|---|---|
| PivotTables | Sales-by-region, top-products, trend pivots |
| XLOOKUP | Pull product name/category/cost from a master Product Table into transaction data |
| VLOOKUP | Legacy-compatible lookup version (some employers still test VLOOKUP specifically) — use it somewhere separate from XLOOKUP so both are visibly demonstrated |
| INDEX-MATCH | Two-way lookup — e.g., pull price for a Product × Region combination from a matrix table |
| SUMIFS | Conditional totals — sales by region AND month, revenue by category AND above a price threshold |
| Conditional Formatting | Low stock flags, negative margin flags, top-performer highlighting |
| Data Validation | Dropdown lists for region/category entry on a "new sale" input form, restrict date entries, prevent negative quantities |
| VBA / Macros | Data cleaning, pivot refresh, report generation, form automation |

### 3. Dataset & structure
- Same retail dataset as before, split into weekly batches
- **Add a second table:** a Product Master List (Product ID, Name, Category, Cost, Reorder Threshold) — this is what makes XLOOKUP/VLOOKUP/INDEX-MATCH necessary, since transaction data will only have Product IDs
- **Add a third table:** a simple Inventory/Stock sheet (Product ID, Current Stock) — needed for the low-stock conditional formatting logic

### 4. Sheet-by-sheet build plan

**Sheet 1 — Raw Data (Intake)**
- Weekly transaction dumps land here, deliberately messy in places (duplicates, blank fields, inconsistent formatting)

**Sheet 2 — Product Master**
- Product ID, Name, Category, Cost, Reorder Threshold
- This is your lookup source table

**Sheet 3 — Inventory**
- Product ID, Current Stock, Last Restocked Date

**Sheet 4 — Clean/Working Data**
- VBA-cleaned version of Raw Data
- XLOOKUP columns pulling Product Name + Category from Product Master
- A second lookup column done in VLOOKUP (e.g., pulling Cost) — explicitly to show you know both
- An INDEX-MATCH column for a two-way lookup (e.g., price matrix by Product × Region, if you build one, or Reorder Threshold by Product+Category combo)

**Sheet 5 — Data Entry Form**
- A simple form area for adding a new manual sale record
- Data Validation: dropdown for Region and Category (sourced from Product Master), restricted date picker, quantity must be positive

**Sheet 6 — Reports (Pivot Dashboard)**
- PivotTables: Sales by Region, Top 10 Products, Monthly Trend
- SUMIFS-based KPI cards above the pivots (e.g., "Total Feb Sales in North Region" as a standalone formula, not a pivot — to explicitly show SUMIFS separate from pivot logic)

**Sheet 7 — Inventory Alerts**
- Conditional formatting: red if Current Stock < Reorder Threshold (pulled via lookup), amber if within 20% of threshold
- Conditional formatting on margin: flag any product where (Price − Cost) is negative or below a target %

**Sheet 8 — Cleaning Log**
- VBA-generated log: rows removed, duplicates found, blanks flagged, timestamp of last run

**Sheet 9 — Summary/Export**
- One-click formatted report combining KPIs + charts, ready to export as PDF

### 5. VBA macros to build
1. `CleanData` — dedupe, fix formats, flag/remove invalid rows, write to Cleaning Log
2. `RefreshLookups` — reapply XLOOKUP/VLOOKUP/INDEX-MATCH formulas to newly cleaned data (in case row counts changed)
3. `RefreshPivots` — refresh all PivotTables/PivotCaches
4. `ApplyFormatting` — reapply conditional formatting rules programmatically (useful if new rows were added)
5. `GenerateSummary` — build/refresh the Summary sheet
6. `ExportToPDF` — export Summary sheet as PDF, named with current date
7. `RunFullReport` (master macro) — calls all of the above in sequence; this is your "one-click" button

### 6. Before/after proof
- Time the full manual process once (clean, lookup, pivot, format, export) vs. one click of `RunFullReport`
- Screenshot/record both

### 7. Deliverables
- Working `.xlsm` file, all macros commented
- README: business problem, architecture (which sheet feeds which), skills demonstrated table (same as section 2 above — recruiters skim these), before/after timing
- Screenshots: cleaning log, lookup columns in action, pivot dashboard, inventory alert formatting, data validation dropdown, exported PDF

### 8. Resume bullet
> "Built a VBA-automated Excel reporting system integrating XLOOKUP, VLOOKUP, INDEX-MATCH, SUMIFS, PivotTables, and Data Validation across a multi-sheet workbook — reduced manual weekly sales/inventory reporting from ~2 hours to under 5 minutes."

### 9. Time estimate
5-6 days at this scope (up from 3-4) — it's genuinely a bigger build now, but it hits every Excel skill on your resume in one coherent project instead of scattered claims.
