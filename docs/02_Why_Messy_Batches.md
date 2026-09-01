# Why We Created Messy Batch Files

## 1. Why batches at all (not one single file)?
The project's entire premise is: "a business receives a new sales export
on a recurring basis (weekly/monthly), and someone manually re-does the
same cleaning + reporting work every time." That repetition is *why*
automation has value.

If we used one single clean file, there would be nothing recurring to
automate — no "before" grind to fix. Splitting the real data into 4
monthly batches (Jan/Feb/Mar/Apr 2017) simulates 4 separate export
deliveries arriving over time, the way a real retail chain would receive
them.

## 2. Why only some batches are messy (2 of 4)?
Real exports aren't uniformly bad. Some arrive clean, some don't (a
system glitch, manual entry error, etc.). A mix is more believable than
either "everything is perfect" or "everything is broken," and it gives
the cleaning macro a genuine before/after contrast to prove against.

- `export_batch_01_2017_clean.xlsx` — clean
- `export_batch_02_2017_MESSY.xlsx` — messy
- `export_batch_03_2017_clean.xlsx` — clean
- `export_batch_04_2017_MESSY.xlsx` — messy

## 3. What "messy" means, specifically
Introduced only into the two MESSY files, via `make_messy_batches.py`:
- **Duplicate rows** (~3% of rows repeated) — tests dedupe logic
- **Inconsistent MonthYear formats** (e.g. "10.2016" vs "2016-10" vs
  "Oct-16" on the same column) — tests date standardization
- **Blank required fields** (~5% of rows missing StoreID or Turnover) —
  tests invalid-row flagging
- **Inconsistent text casing/whitespace** in Dept. Name / City (~20% of
  rows, e.g. "DRY", "dry", "  Dry  ") — tests text normalization

All of this was generated with `random.seed(42)`, so it's reproducible —
running the script again produces identical files, nothing is
hand-picked or hidden.

## 4. What `make_messy_batches.py` is
A one-time Python helper script, NOT part of the Excel/VBA deliverable.
Its only job was to produce the 4 batch files above from the original
data. It has already been run; you don't need to run it again unless you
want to regenerate or tweak the messiness rules.

## 5. What happens to these 4 files next
They get imported/combined into a single **`Raw_Data` sheet** inside the
real `.xlsm` workbook — as if all 4 months' exports had been pasted in
over time. That combined `Raw_Data` sheet is what the VBA `CleanData`
macro will actually run against.

Sequence from here:
1. Build the `.xlsm` workbook shell (sheets, no macros yet)
2. Import all 4 batches into `Raw_Data` (combined, messiness intact)
3. Write the `CleanData` macro to fix exactly the issues listed in
   section 3 above
4. Prove it worked: before/after row counts, before/after screenshots
