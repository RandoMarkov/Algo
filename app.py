"""xlwings button layer — the ONLY Windows/Excel-specific module.

Each function is bound to one button in the front-end workbook. They are thin:
read a few control cells, call the (platform-independent, already-tested)
backend in ingest.py / compute.py, and write results back into the live book.

Wire-up in the workbook's VBA (one line per button), e.g.:

    Sub Btn_RecordFund()
        RunPython "import app; app.record_fund()"
    End Sub
    Sub Btn_Compute()
        RunPython "import app; app.compute_portfolio()"
    End Sub
    Sub Btn_Charts()
        RunPython "import app; app.make_charts()"
    End Sub
"""
from pathlib import Path

import xlwings as xw

import config
import db
import ingest
import compute


# --- helpers ----------------------------------------------------------------

def _book():
    """The workbook that invoked the macro."""
    return xw.Book.caller()


def _db_for(book):
    """Use a funds.db sitting next to the front-end workbook."""
    return db.connect(Path(book.fullname).parent / "funds.db")


def _ctrl(book, key):
    return book.sheets[config.CTRL_SHEET].range(config.CTRL_CELLS[key]).value


def _status(book, msg):
    """Write a status line the user can see (Control status cell) + a popup."""
    book.sheets[config.CTRL_SHEET].range(config.STATUS_CELL).value = msg
    try:
        book.app.alert(msg, "Fund tool")
    except Exception:
        pass


def _as_iso(value):
    """xlwings hands dates back as datetime; normalise to ISO string."""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


# --- BUTTON 1: record a fund file into the DB -------------------------------

def record_fund():
    book = _book()
    name = _ctrl(book, "fund_file")
    if not name:
        _status(book, "Put the fund file name in Control!B2 first.")
        return
    path = Path(name)
    if not path.is_absolute():                       # same directory as workbook
        path = Path(book.fullname).parent / name
    if not path.exists():
        _status(book, f"File not found: {path}")
        return

    conn = _db_for(book)
    try:
        s = ingest.ingest_fund_file(path, conn)
    finally:
        conn.close()
    _status(book, f"Recorded {s['fund_id']} @ {s['as_of_date']}: "
                  f"{s['n_holdings']} holdings, {s['n_geography']} geo, "
                  f"{s['n_currency']} ccy, {s['n_metrics']} metrics.")


# --- BUTTON 2: compute portfolio look-through -------------------------------

def _sync_portfolio_from_sheet(book, conn, portfolio_id, base_ccy):
    """Read the Portfolio sheet table (Fund ID | Effective Date | Weight) into DB."""
    sht = book.sheets[config.PORTFOLIO_SHEET]
    table = sht.range("A1").expand().value          # incl header row
    if not table or len(table) < 2:
        return 0
    header = [str(h).strip() for h in table[0]]
    ix = {h: header.index(h) for h in ("Fund ID", "Effective Date", "Weight")}
    allocs = []
    for row in table[1:]:
        if row[ix["Fund ID"]] in (None, ""):
            continue
        allocs.append({
            "fund_id": str(row[ix["Fund ID"]]).strip(),
            "effective_date": _as_iso(row[ix["Effective Date"]]),
            "weight": float(row[ix["Weight"]]),
        })
    db.replace_allocations(conn, portfolio_id, portfolio_id, base_ccy, allocs)
    conn.commit()
    return len(allocs)


def _write_results(book, result):
    """Write the computed tables onto the Results sheet (dynamic dimensions)."""
    if config.RESULTS_SHEET in [s.name for s in book.sheets]:
        book.sheets[config.RESULTS_SHEET].clear()
    else:
        book.sheets.add(config.RESULTS_SHEET, after=book.sheets[-1])
    res = book.sheets[config.RESULTS_SHEET]

    res.range("A1").value = [["Portfolio", result["portfolio_id"]],
                             ["As-of date", result["as_of_date"]],
                             ["Invested weight", round(result["invested_weight"], 6)],
                             ["Cash residual", round(result["cash_residual"], 6)]]
    row = 6

    def section(title, df):
        nonlocal row
        if df.empty:
            return
        res.range((row, 1)).value = title
        res.range((row + 1, 1)).value = [list(df.columns)] + df.values.tolist()
        row += 1 + len(df) + 2

    # one section per exposure dimension, then holdings, then metrics
    for dim in sorted(result["exposures"]):
        df = result["exposures"][dim]
        df = df[df["weight"].abs() > 1e-9]
        label = config.DIM_LABELS.get(dim, dim)
        cov = result["coverage"].get(dim, 0.0)
        section(f"{label} (coverage {cov:.0%})", df)
    hd = result["holdings"]
    section("Top holdings (look-through)", hd[hd["weight"].abs() > 1e-9].head(15))
    section("Weighted metrics", result["metrics"])


def compute_portfolio():
    book = _book()
    pid = _ctrl(book, "portfolio_name")
    date = _as_iso(_ctrl(book, "as_of_date"))
    if not pid or not date:
        _status(book, "Set portfolio name (B3) and as-of date (B4) in Control.")
        return

    conn = _db_for(book)
    try:
        n = _sync_portfolio_from_sheet(book, conn, pid, "EUR")
        result = compute.compute_portfolio(conn, pid, date)
    finally:
        conn.close()

    _write_results(book, result)
    _status(book, f"Computed {pid} as-of {date} "
                  f"({n} allocations; invested {result['invested_weight']:.1%}). "
                  f"Now click 'Make charts'.")


# --- BUTTON 4: exposure evolution over a date range -------------------------

def evolution():
    book = _book()
    pid = _ctrl(book, "portfolio_name")
    dim = _ctrl(book, "evo_dimension") or "geography_country"
    start, end = _as_iso(_ctrl(book, "evo_start")), _as_iso(_ctrl(book, "evo_end"))
    step = int(_ctrl(book, "evo_step_months") or 1)
    if not pid or not start or not end:
        _status(book, "Set portfolio (B3), dimension (B5), start (B6), end (B7).")
        return

    conn = _db_for(book)
    try:
        _sync_portfolio_from_sheet(book, conn, pid, "EUR")   # ensure DB is current
        dates = compute.date_steps(start, end, step)
        evo = compute.compute_evolution(conn, pid, dim, dates)
    finally:
        conn.close()
    if evo.empty:
        _status(book, f"No '{dim}' data for {pid} in that range.")
        return
    evo = compute.top_columns(evo, 8)

    # write table + line chart onto the Evolution sheet
    if config.EVOLUTION_SHEET in [s.name for s in book.sheets]:
        ev = book.sheets[config.EVOLUTION_SHEET]
        ev.clear()
        for c in list(ev.charts):
            c.delete()
    else:
        ev = book.sheets.add(config.EVOLUTION_SHEET, after=book.sheets[-1])

    ev.range("A1").value = ([["date"] + [str(c) for c in evo.columns]]
                            + [[str(idx)] + [float(v) for v in row]
                               for idx, row in evo.iterrows()])
    nrow, ncol = 1 + len(evo), 1 + evo.shape[1]
    src = ev.range((1, 1), (nrow, ncol))
    label = config.DIM_LABELS.get(dim, dim)
    top = (nrow + 2) * 16
    for kind, dx in (("line", 0), ("area_stacked", 640)):
        chart = ev.charts.add(left=10 + dx, top=top, width=620, height=320)
        chart.set_source_data(src)
        chart.chart_type = kind
        chart.api[1].HasTitle = True
        chart.api[1].ChartTitle.Text = f"{label} evolution ({kind})"
    _status(book, f"Evolution of '{dim}' for {pid}: {len(dates)} dates "
                  f"{dates[0]}…{dates[-1]} (line + stacked area) on the Evolution sheet.")


# --- BUTTON 3: draw charts from the Results sheet ---------------------------

def make_charts():
    """Draw native Excel charts on the Charts sheet from the Results tables."""
    book = _book()
    if config.RESULTS_SHEET not in [s.name for s in book.sheets]:
        _status(book, "Run 'Compute portfolio' first.")
        return
    res = book.sheets[config.RESULTS_SHEET]

    if config.CHARTS_SHEET in [s.name for s in book.sheets]:
        ch = book.sheets[config.CHARTS_SHEET]
        for c in list(ch.charts):
            c.delete()
    else:
        ch = book.sheets.add(config.CHARTS_SHEET, after=book.sheets[-1])

    # Re-discover sections by scanning column A: a section is a TITLE row
    # (col A text, col B empty) followed by a header row (col A & B both text)
    # and data rows until the next blank. Works for any set of dimensions.
    PIE_MAX = 8
    used = res.range("A1").end("down").row
    top = 20
    r = 6                                            # data sections start at row 6
    while r <= used:
        a, b = res.range((r, 1)).value, res.range((r, 2)).value
        if a and not b:                              # title row
            title = str(a)
            hdr = r + 1
            last = hdr
            while res.range((last + 1, 1)).value not in (None, ""):
                last += 1
            n = last - hdr                           # data rows
            if not title.startswith("Weighted metrics") and n >= 1:
                chart = ch.charts.add(left=10, top=top, width=380, height=230)
                chart.set_source_data(res.range((hdr, 1), (last, 2)))
                chart.chart_type = "pie" if n <= PIE_MAX else "bar_clustered"
                chart.api[1].HasTitle = True
                chart.api[1].ChartTitle.Text = title
                top += 250
            r = last + 1
        else:
            r += 1

    _status(book, "Charts updated on the Charts sheet.")
