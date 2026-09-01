# Notes: What `build_workbook_shell.py` Does

A one-time Python helper script (not part of the Excel/VBA deliverable)
that assembled the main project workbook: `RetailSalesWorkloadAutomation.xlsx`.

## Step by step

1. **Created a new blank workbook.**

2. **Built the `Raw_Data` sheet:**
   - Opened all 4 batch files (`export_batch_01_2017_clean.xlsx` through
     `export_batch_04_2017_MESSY.xlsx`)
   - Took the header row once
   - Stacked all their data rows together, one after another
   - Wrote the combined result (3,450 rows) into `Raw_Data`
   - Messiness from batches 2 and 4 (MESSY) is preserved as-is — nothing
     was cleaned at this stage

3. **Built the `Store_Master` sheet:**
   - Opened the original file (`salesworkload_ORIGINAL.xlsx`)
   - Read the `opening_schemes` sheet
   - Pulled out just 4 columns: `id`, `Store name`, `Region`, `Scheme`
     (dropped the monthly hours grid — not needed for lookups)
   - Wrote that as a clean 50-row store list

4. **Created 6 more sheets, left empty (placeholders for later stages):**
   - `Clean_Data`
   - `Data_Entry_Form`
   - `Reports_Pivot`
   - `Alerts`
   - `Cleaning_Log`
   - `Summary_Export`

5. **Saved everything as `RetailSalesWorkloadAutomation.xlsx`.**

## Result

One workbook, 8 sheets total:

| Sheet | State |
|---|---|
| Raw_Data | 3,450 rows, messy (before-state) |
| Store_Master | 50 rows, clean, ready for lookups |
| Clean_Data | empty — next stage builds this |
| Data_Entry_Form | empty |
| Reports_Pivot | empty |
| Alerts | empty |
| Cleaning_Log | empty |
| Summary_Export | empty |

## Status
Script has already been run. Its job is done — the output file
(`RetailSalesWorkloadAutomation.xlsx`) is what we keep working in from
here. No need to re-run this script unless the batch files or
Store_Master source data change.
