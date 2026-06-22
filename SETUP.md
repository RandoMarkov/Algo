# Setup — turning `PortfolioTool.xlsx` into the working `.xlsm` (Windows)

`PortfolioTool.xlsx` is fully built (sheets, inputs, sample portfolio,
`xlwings.conf`). These one-time steps add the Python link and the 3 buttons.
Everything after that is just clicking buttons.

## 0. Put the project on the Windows PC
Copy this whole folder somewhere local, e.g. `C:\Tools\fund_portfolio\`.
Keep the fund data files (`*.xlsx`) in this same folder so button ① can find
them by name.

## 1. Python + libraries (one time)
```powershell
cd C:\Tools\fund_portfolio
python -m venv .venv
.venv\Scripts\activate
pip install pandas openpyxl xlwings
xlwings addin install        # installs the Excel ribbon add-in
```

## 2. Open the workbook and point xlwings at this Python
1. Open `PortfolioTool.xlsx`.
2. Go to the **xlwings.conf** sheet. `Interpreter_Win` already points at
   `.venv\Scripts\pythonw.exe` relative to the workbook and `PYTHONPATH` at the
   workbook folder — adjust only if you used a different venv.

## 3. Add the macros (one time)
1. Press **Alt+F11** (VBA editor).
2. **File ▸ Import File…** → choose `FundTool.bas` from this folder.
3. **Tools ▸ References…** → tick **xlwings** → OK. (This makes `RunPython`
   available; the add-in from step 1 provides it.)
4. Close the editor.

## 4. Draw the four buttons (one time)
On the **Control** sheet, for each of the four:
1. **Insert ▸ Shapes** (or **Developer ▸ Insert ▸ Button**) → draw a button.
2. Label it (Record fund / Compute / Make charts / Evolution).
3. Right-click ▸ **Assign Macro…** → pick `Btn_RecordFund`, `Btn_Compute`,
   `Btn_Charts`, `Btn_Evolution` respectively.

## 5. Save as macro-enabled
**File ▸ Save As ▸ Excel Macro-Enabled Workbook (\*.xlsm)** →
`PortfolioTool.xlsm`. Done.

---

## Daily use
1. **Record a fund**: type a fund file name into **Control!B2** → click ①.
   Repeat for each fund/date you receive. (Re-recording the same fund+date just
   replaces it.)
2. **Build the portfolio**: on the **Portfolio** sheet list
   `Fund ID | Effective Date | Weight` (add a new dated row to rebalance —
   weights may vary over time).
3. **Compute**: set **Control!B3** (portfolio name) and **B4** (as-of date) →
   click ② → numbers land on **Results** (each exposure tagged with its
   coverage %).
4. **Charts**: click ③ → pie/bar charts land on **Charts**.
5. **Evolution** (optional): set **B5** (dimension — pick from the dropdown),
   **B6/B7** (start/end date), **B8** (step in months) → click ④ → a
   time-series table plus a **line chart** (per-bucket trend) and a
   **stacked-area chart** (composition) land on **Evolution**.

Status messages and any errors show in **Control!B10**.

## Troubleshooting
- *"RunPython not found"* → step 3.3 (References ▸ xlwings) was skipped.
- *Import errors / module not found* → `PYTHONPATH` in `xlwings.conf` must be
  the folder holding `app.py`; the venv must have pandas/openpyxl installed.
- *Nothing happens / wrong Python* → check `Interpreter_Win` in `xlwings.conf`.
- Prefer no VBA at all for a quick test? With the venv active you can run the
  whole pipeline headless: `python test_pipeline.py`.
