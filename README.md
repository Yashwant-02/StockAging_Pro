# Stock Aging & Refill Analytics (Simplified & Fast)

Matches **Outward + Sales + Current Stock** files by **Barcode + Store** and produces
exactly **3 Excel sheets** — no FIFO, no extra sheets, no row-by-row loops.

## Setup

```bash
python setup_project.py
streamlit run app.py
```

(or manually: `pip install -r requirements.txt` then `streamlit run app.py`)

## How it works

1. Upload the **Outward**, **Sales** and **Current Stock** files (xlsx/xls/csv).
   Optionally upload a 4th **WH Stock** file (central warehouse stock).
2. Map each file's columns (Store, Barcode, Qty, Date, Article/Color/Size).
   The WH file maps by **Barcode only** (no Store, since WH stock is central).
3. Adjust **Fresh Cutoff** / **RTV Cutoff** (days) in the sidebar if needed
   (defaults: 20 / 40).
4. Click **Generate Report**.
5. Download the Excel file with 3 sheets.

## Core rules

- **Current Stock file is the only source of truth for current stock.**
  Stock is never calculated as `Outward − Sales`.
- **Sales that happened before a Barcode+Store's first tracked Outward Date
  are ignored everywhere** (Performance, Refilling, RTV_IST). Those sales
  belong to older stock outside the Outward file's date window, so counting
  them would inflate Sold %/Selling Days with unrelated history.

## Output sheets

### 1. Performance
Only Barcode+Store combinations that appear in **both** Outward and (cleaned)
Sales. `Selling Days = Sale Date − Outward Date`, tagged:

| Selling Days | Tag |
|---|---|
| 0–15 | Fast |
| 16–30 | Good |
| 31–60 | Moderate |
| 61–90 | Slow |
| 91+ | Very Slow |

### 2. Refilling
Based on the Current Stock lifecycle for every Barcode+Store that currently
has stock and/or was ever outward'd. Single `Action` column
(`Fresh Cutoff` / `RTV Cutoff` are set in the sidebar, default 20 / 40 days):

- **No Outward record found at all** → `RTV/IST` (Age Days defaults to RTV Cutoff)
- **Age 0 – Fresh Cutoff** → `Fresh`
- **Age Fresh Cutoff – RTV Cutoff** → `Monitor`
- **Age > RTV Cutoff, never sold (Sales Qty = 0)** → `RTV/IST`
- **Age > RTV Cutoff, Sales Qty ≥ Outward Qty but stock still > 0** → `RTV/IST`
  (the tracked outward is fully sold, so whatever is still sitting in stock
  is untracked/older stock — recommend returning it)
- **Age > RTV Cutoff, partially sold, stock still remaining** → `Monitor`
- **Age > RTV Cutoff, sold out completely, sold fast (Selling Days < RTV Cutoff)**
  → checked against the WH file:
  - WH has stock for that Barcode → `Refill`
  - WH has no stock (or no WH file uploaded, or Barcode not found) → `Not in WH`
    (falls back to `Refill` automatically if no WH file was uploaded)
- **Age > RTV Cutoff, sold out completely, sold slow (Selling Days ≥ RTV Cutoff)**
  → `Sold – No Refill`

`Age Days` is measured from the **Last Outward Date** to the Report Date
(set in the sidebar; used only for this calculation, never shown as a column).
`Sold %` and `Remaining Stock %` are capped at 100% for readability.

### 3. RTV_IST
`Outward Qty > 0 AND Sales Qty = 0 AND Current Stock = 0` → `Status = "RTV / IST"`

## Performance / speed

- Every file is cleaned and aggregated **once** with `groupby` (no per-unit FIFO
  allocation, no `iterrows()`).
- Each sheet is built with a single set of `merge()` calls.
- A built-in validation check flags a sheet if aggregation ever produces more
  rows than unique Barcode+Store combinations (a sign the Store/Barcode
  mapping is wrong).
- All date columns are exported as `dd-mm-yyyy`.


"""
Stock Aging & Refill Analytics (Simplified & Fast)
----------------------------------------------------
Matches Outward, Sales, Current Stock and (optionally) a central Warehouse (WH)
stock file by Barcode / Barcode+Store to produce EXACTLY 3 business-friendly
Excel sheets:

  1. Performance  - how fast SOLD quantity sold (Outward Date -> Sale Date)
  2. Refilling    - current-stock lifecycle & refill decision
  3. RTV_IST      - outward stock untraceable in both Sales and Current Stock

Design rules (locked with business):
  - CURRENT STOCK FILE is the only source of truth for current stock.
    Stock is NEVER derived as Outward - Sales.
  - No FIFO / batch allocation, no row-wise loops. Every file is cleaned and
    aggregated exactly once with pandas groupby, then merged once per sheet.
  - No extra/unnecessary sheets (no FIFO detail, no exceptions sheet, no
    reconciliation sheet) - only the 3 sheets above are exported.
  - Report Date is used ONLY internally to compute Age Days. It is never
    shown as a column in the output.
  - All date columns are exported in dd-mm-yyyy format, and shown the same
    way (date only, no time) in the on-screen preview tables.
  - Sales that happened BEFORE a Barcode+Store's first tracked Outward Date
    are excluded from every calculation (that sale belongs to older,
    untracked stock outside the Outward file's window).
  - Warehouse (WH) file is optional. It is matched by Barcode ONLY (no Store,
    since WH stock is central/shared across stores). It is only consulted to
    confirm a "Refill" candidate; every other Action is unaffected by it.

Implementation notes:
  - The full clean -> aggregate -> merge -> Excel-build pipeline runs ONLY
    when "Generate Report" is clicked. Results (dataframes + the built Excel
    bytes) are cached in st.session_state, so clicking the download button
    (which triggers a Streamlit rerun) reuses the cached results instead of
    recomputing/rebuilding everything from scratch.
  - xlsxwriter's "constant_memory" streaming mode is intentionally NOT used:
    it flushes rows before column formats (like the dd-mm-yyyy date format)
    are applied, which silently corrupts/blanks already-flushed rows. Our
    aggregated data is small enough that normal (non-streaming) mode is fine.
"""