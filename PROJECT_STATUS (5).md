# Project Status: Retail Sales & Workload Reporting Automation
Last updated: this session. Read this first if resuming later.

## IMPORTANT: File location note
Your live working file is the `.xlsm` you saved yourself in Excel
(with VBA macros in it) — NOT the `.xlsx` files listed below, which are
earlier snapshots without macros. Use your own saved `.xlsm` going forward.

---

## ✅ DONE SO FAR

### 1. Planning docs
- `00_Objectives.md` — business problem, scope, target resume bullet
- `01_Requirements.md` — full spec: sheets, macros, skills-map, using your
  REAL columns (MonthYear, StoreID, Dept. Name, Turnover, etc.)
- `02_Why_Messy_Batches.md` — explains why we simulate messy monthly
  exports and what happens to them

### 2. Source data understood
- `salesworkload_ORIGINAL.xlsx` — your real uploaded file, 2 sheets:
  - `sales_figures` (~7,650 usable rows after excluding 8 pre-existing
    blank/footer artifact rows)
  - `opening_schemes` (50 real stores + monthly hours-type grid)

### 3. Simulated "before" state (messy batches)
Script `make_messy_batches.py` split real data into 4 monthly batches
(Jan–Apr 2017) and deliberately broke 2 of them:
- `export_batch_01_2017_clean.xlsx` — clean
- `export_batch_02_2017_MESSY.xlsx` — messy (dupes, bad dates, blanks, casing)
- `export_batch_03_2017_clean.xlsx` — clean
- `export_batch_04_2017_MESSY.xlsx` — messy (same issues)

### 4. Workbook shell built
Script `build_workbook_shell.py` combined the 4 batches into one workbook
with 8 sheets: `Raw_Data` (3,450 rows, messy), `Store_Master` (50 stores),
plus 6 empty placeholder sheets (`Clean_Data`, `Data_Entry_Form`,
`Reports_Pivot`, `Alerts`, `Cleaning_Log`, `Summary_Export`).
Saved as `RetailSalesWorkloadAutomation.xlsx`.

### 5. VBA macros written (in your Excel file, in a Module)
- **`CleanData`** — dedupes exact-match rows, standardizes MonthYear format,
  trims/proper-cases Dept. Name & City, flags rows with blank
  StoreID/Turnover as "INVALID" (kept, not deleted), writes results +
  timestamp to `Cleaning_Log`.
  - Verified working: In 3450 → Duplicates removed 24 → Flagged invalid 86
    → Out 3426 (this run is now being redone with 2 more fixes, see below)
- **`RefreshLookups`** — called automatically at the end of `CleanData`.
  Writes 3 lookup columns into `Clean_Data`:
  - Column P: Store Name — auto-detects if XLOOKUP is supported (365/2021);
    falls back to INDEX-MATCH if not (confirmed: your Excel 2019 → uses
    INDEX-MATCH fallback, labeled `Store_Name_INDEXMATCH`)
  - Column Q: Region — via VLOOKUP
  - Column R: Scheme — via INDEX-MATCH
  - Verified working: all 3 columns populated with real data

### 6. Bugs found and fixed mid-build (real data-quality lessons — good
   interview material)
- **Bug 1 (fixed):** `StandardizeMonthYear` failed on values like "Feb-17"
  because Excel silently auto-converts such text into a real Date/serial
  number when written to a cell — VBA was reading a mangled date instead
  of the expected string. First fix attempt used `IsDate()`, which turned
  out to be insufficient (see Bug 3).
- **Bug 2 (fixed):** the dataset has a `Dept. Name = "all"` value
  (450 rows) that is a store-level ROLLUP/grand-total row, not a real
  department — its turnover (~7.1B) was polluting department-level and
  month-level pivots. Fixed by excluding these rows entirely inside
  `CleanData` (not just filtering at report time), with the excluded
  count logged separately in `Cleaning_Log` (`rollupCount`).
- **Bug 3 (fixed):** even after Bug 1's fix, values like "Apr-17" were
  still corrupting to "04.2026" instead of "04.2017". Root cause: VBA's
  `IsDate("Apr-17")` returns True but interprets it as "April 17th of the
  CURRENT system year" (day=17, not year=2017) — a genuine locale/parsing
  ambiguity, not a typo. Fixed by (a) checking explicit text patterns
  (`MM.YYYY`, `YYYY-MM`, `Mon-YY`) BEFORE any date-type fallback, and
  (b) replacing `IsDate()` with `VarType(val) = vbDate`, which only
  catches values that are truly Date-typed, not text that merely looks
  date-like. Also added `NumberFormat = "@"` before writing the cleaned
  MonthYear value, forcing Text format so Excel can't re-convert it later.
  See `04_Note_MonthYear_Text_Format.md` for why staying as text is
  correct (not a remaining issue).

### 7. PivotTables built manually on `Reports_Pivot` (by you, in Excel)
— all 3 rebuilt and verified clean after Bug 2 + Bug 3 fixes:
- Pivot #1 (cell A3 area): Turnover by Country — done, verified
- Pivot #2 (cell D3 area): Turnover by Dept. Name, sorted largest→smallest
  — done, verified, "all" rollup no longer appears
- Pivot #3 (cell G3 area): Turnover by MonthYear (month trend) — done,
  verified: clean 4 rows (01.2017–04.2017), no garbled dates, Grand Total
  8,514,564,405

---

## ✅ Stages 1–7 above are now fully complete and verified.

## 🔲 REMAINING STAGES (not started yet)

5. **SUMIFS KPI cards** — DONE. On `Reports_Pivot`, 3 standalone formulas
   below the pivots (rows 24-27): Total UK Turnover (all months) =
   1,009,682,868; Total Germany Turnover, April 2017 = 353,679,555; Total
   Food Dept Turnover, April 2017 = 698,290,866. Note: initially tried a
   dynamic "latest month" via MAX() on the MonthYear column, but MAX()
   can't read text values (returned 0) — fixed by hardcoding "04.2017"
   explicitly in both the label and formula, since MonthYear is
   intentionally text (see companion doc).
6. **Conditional Formatting on `Alerts` sheet** — DONE, via VBA macro
   `ApplyFormatting` (not manual formatting) so it's robust to row-count
   changes. The macro copies `Clean_Data` into `Alerts` as values (auto-
   sizing to current row count), clears old formatting, then applies 3
   rules: Turnover < 500,000 → orange cell highlight; Customer = 0 →
   yellow cell highlight; Is_Valid = "INVALID" → light red highlight on
   the whole row. Verified working.
7. **Data Validation on `Data_Entry_Form`** — DONE. Headers set up
   (MonthYear, Country, StoreID, City, Dept. Name, Sales units, Turnover).
   Country dropdown (List, sourced from `Store_Master!$C$2:$C$51`).
   Dept. Name dropdown (List, sourced from `Clean_Data!$G$2:$G$3227` —
   note: shows duplicates since it pulls every row, not a deduped list;
   a normal simplification, worth noting honestly in README rather than
   over-engineering with a helper column). Sales units and Turnover both
   restricted to positive numbers (Whole number / Decimal, greater than 0).
8. **`RefreshPivots` macro** — refreshes all 3 PivotCaches with one click
   (not rebuild from scratch — per your earlier correct decision to build
   pivots manually once, then just refresh them)
9. ~~**`ApplyFormatting` macro**~~ — DONE, see item 6 above (built ahead
   of sequence since it fit naturally with the Alerts sheet work)
10. **`GenerateSummary` macro** — DONE. Builds `Summary_Export`: title +
    timestamp, "Key Figures" section (3 SUMIFS KPIs copied as values from
    Reports_Pivot), then all 3 pivot snapshots (Country, Dept, Month)
    copied as values side-by-side in a dashboard layout. Verified working
    on first run, clean output, no fixes needed.
11. **`ExportToPDF` macro** — DONE. Exports `Summary_Export` as a dated
    PDF (e.g. `Sales_Summary_2026-09-01.pdf`) in the workbook's folder.
    Hit real layout issues along the way (worth noting for
    README/interview): first export broke the Month-over-Month table
    layout; tried fixing via VBA PageSetup properties
    (FitToPagesWide/Tall), which over-corrected and cut off all KPI/pivot
    VALUES (only labels remained). Fixed by reverting to the simple
    macro (just `ExportAsFixedFormat`, no PageSetup code) and instead
    setting Landscape orientation + Print Area + "Fit to 1 page wide"
    manually, once, directly in Excel's Page Layout tab (settings persist
    with the saved workbook). Lesson: a one-time visual/manual setting
    beat repeated blind VBA property tweaking here. Verified working:
    clean single-page PDF, all 3 tables + KPIs aligned correctly.
12. **`RunFullReport` master macro** — DONE. Chains CleanData →
    RefreshPivots → ApplyFormatting → GenerateSummary → ExportToPDF in
    one click. Added a module-level `silentMode` flag so intermediate
    macros suppress their own MsgBox popups when called from
    RunFullReport (only the final timing summary shows) — required
    wrapping the MsgBox line in ALL 6 subs (CleanData, RefreshLookups,
    ApplyFormatting, RefreshPivots, GenerateSummary, ExportToPDF) with
    `If Not silentMode Then ... End If`; initially only CleanData was
    wrapped, causing all popups to still appear — fixed once all 6 were
    updated. **Verified working: real automated time = 11.8 seconds.**
13. **Before/after timing proof** — DONE. Automated: 11.8 seconds
    (verified via RunFullReport's own popup). Manual baseline: ~1 hour
    (reasoned estimate, not stopwatch-measured — documented honestly as
    such). Reasoning: this is a RECURRING task — every time a new
    month's dataset arrives, someone would redo cleaning, 3 lookup types
    across ~3,000 rows, 3 pivots, conditional formatting, the entry form,
    and the summary/PDF export by hand, each cycle. Automating it turns
    that recurring ~1 hour into one 11.8-second click.
14. **README.md** — business problem, architecture, skills-demonstrated
    table, before/after timing, screenshots list
15. **Screenshots** — Cleaning_Log, lookup columns, pivot dashboard,
    conditional formatting, data validation dropdown, exported PDF

## Target Resume Bullet (FINALIZED)
> "Built a VBA-automated Excel reporting system integrating XLOOKUP/
> INDEX-MATCH, VLOOKUP, SUMIFS, PivotTables, and Data Validation across a
> multi-sheet retail sales workbook — reduced recurring manual reporting
> time from ~1 hour to under 12 seconds with a one-click macro."

## How to resume this conversation later
Paste this file's contents into a new chat, or re-upload the file, and say
"continue from REMAINING STAGES, step 14 (README.md)" (or wherever you
actually left off — update the checkboxes above as you go).

## Companion docs
- `04_Note_MonthYear_Text_Format.md` — explains why MonthYear is
  deliberately kept as text (not converted to a real date), and the
  IsDate/VarType bug that led to this decision.
