# Fund & Portfolio look-through tool

A light SQLite-backed tool. Fund data arrives as Excel files (4 sheets each);
a front-end Excel workbook with **3 buttons** drives everything, and all real
logic is Python (minimal VBA).

## What each button does
1. **Record a fund** → reads one fund's `.xlsx` (Synthesis / Geography /
   Currency / Holdings sheets) into `funds.db`.
2. **Compute portfolio** → look-through exposure & holdings of a portfolio
   (a time-varying basket of funds) as-of a chosen date → `Results` sheet.
3. **Make charts** → pie/bar charts of those results → `Charts` sheet.

## Files
| file | role | runs on |
|---|---|---|
| `config.py` | paths + expected sheet/column layout — **edit to match your real files** | any |
| `db.py` | SQLite schema + the "as-of" query helpers | any |
| `ingest.py` | parse one fund workbook → DB | any |
| `compute.py` | portfolio look-through + evolution | any |
| `charts.py` | headless: results+charts → standalone `.xlsx` | any |
| `app.py` | the 3 xlwings button entry points | **Windows + Excel** |
| `make_samples.py` | generate sample fund files | any |
| `test_pipeline.py` | full end-to-end check (no Excel needed) | any |

## Try it without Excel (headless)
```bash
python -m venv .venv && . .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install pandas openpyxl xlwings
python test_pipeline.py        # ingests samples, computes, writes sample_results.xlsx
```

## Three fund types
Files come in three layouts — **Equity / Fixed Income / Alternatives**. They
share a `Synthese` identity sheet and an `Inventaire` holdings sheet, but each
reports different exposure sheets (Equity: region/DM/EM geography, two market-cap
schemes; Fixed Income: rating, maturity, hierarchical sector; Alternatives:
long/short capitalisation). The type is **auto-detected** from the Synthese
`Categorie` cell; `config.PROFILES` declares, per type, which sheets feed which
dimension. The generic `fund_exposures(dimension, bucket, weight)` table stores
all of them, so mixed-type portfolios combine cleanly.

## Key design points
- **As-of semantics.** Funds report at discrete dates with a lag. Every query
  for date *D* uses each fund's *latest snapshot ≤ D* independently
  (`db.latest_date_for`). Staleness is flagged (`config.STALENESS_DAYS`).
- **Coverage.** A dimension only some funds report (e.g. `credit_rating` exists
  only for bonds; bonds have no `currency` sheet) is shown with the % of the
  portfolio that actually reports it — so a 35%-coverage rating breakdown is not
  mistaken for a whole-portfolio view.
- **Multi-sheet dimensions.** Several sheets can feed one dimension (Devise
  DM+EM → `currency`); duplicate buckets across/within sheets (USD in both) are
  summed.
- **Time-varying weights** on both sides: portfolio→fund allocations
  (`portfolio_allocations.effective_date`) and fund→security holdings
  (`fund_holdings.as_of_date`).
- **Cash residual.** Fund weights need not sum to 1; the shortfall is reported
  as `cash_residual` rather than silently normalised.
- **Idempotent ingest.** Re-recording a fund for a date replaces that snapshot.

## Wiring the buttons (Windows, one-time)
1. Install the add-in: `pip install xlwings`, then `xlwings addin install`.
2. New workbook saved as **macro-enabled** `.xlsm` in this folder, with sheets:
   - `Control`: `B2`=fund file name, `B3`=portfolio name, `B4`=as-of date.
   - `Portfolio`: table `Fund ID | Effective Date | Weight` starting at `A1`.
3. In the VBA editor (Alt+F11) add a module with one sub per button:
   ```vba
   Sub Btn_RecordFund(): RunPython "import app; app.record_fund()": End Sub
   Sub Btn_Compute():    RunPython "import app; app.compute_portfolio()": End Sub
   Sub Btn_Charts():     RunPython "import app; app.make_charts()": End Sub
   ```
   In the xlwings ribbon set **Interpreter** to this folder's `.venv` and
   leave PYTHONPATH pointing at this folder. Insert 3 shapes/buttons and
   assign each macro. That VBA is the *only* VBA in the project.

## Adapting to your real fund files
The parser is already wired to the three template layouts in
`incoming_samples/`. If a real file adds/renames a sheet, edit that type's
profile in `config.PROFILES` — e.g. add `_ex("NewSheet", "new_dimension")`, or
adjust `weight_col`/`label_cols`/`ffill_cols` for an oddly-shaped sheet. No
changes to `ingest.py` needed. `make_samples.py` shows how dated snapshots are
produced from the templates for the as-of/evolution demo.
