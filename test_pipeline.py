"""End-to-end proof on the REAL template layouts (no Excel required).

Makes dated copies of the 3 templates, ingests them (auto-detecting type),
builds a mixed-type portfolio with time-varying weights, computes look-through
exposure/holdings as-of two dates, and renders charts.
Run: python test_pipeline.py
"""
from pathlib import Path

import make_samples
import ingest
import compute
import charts
import db

BASE = Path(__file__).resolve().parent
DB = BASE / "test_funds.db"

# fund ids are the ISINs inside the templates
EQ, FI, ALT = "LU1670707527", "LU1333337", "LU70707527ALT"


def main():
    if DB.exists():
        DB.unlink()
    make_samples.main()
    conn = db.connect(DB)

    print("\n== Ingest (type auto-detected) ==")
    fund_of = {}
    for f in sorted((BASE / "work_funds").glob("*.xlsx")):
        s = ingest.ingest_fund_file(f, conn)
        fund_of[f.name] = s["fund_id"]
        print(f"  {f.name:22s} -> {s['profile']:13s} {s['fund_id']:14s} "
              f"@ {s['as_of_date']}  {s['n_holdings']} holdings, "
              f"dims={s['dimensions']}")

    # portfolio with TIME-VARYING, MIXED-TYPE weights
    db.replace_allocations(conn, "MULTI", "Multi-Asset", "EUR", [
        {"fund_id": EQ,  "effective_date": "2026-01-01", "weight": 0.50},
        {"fund_id": FI,  "effective_date": "2026-01-01", "weight": 0.50},
        {"fund_id": EQ,  "effective_date": "2026-06-01", "weight": 0.40},
        {"fund_id": FI,  "effective_date": "2026-06-01", "weight": 0.35},
        {"fund_id": ALT, "effective_date": "2026-06-01", "weight": 0.25},
    ])
    conn.commit()

    for date in ("2026-04-15", "2026-06-20"):
        print(f"\n== Portfolio MULTI as-of {date} ==")
        res = compute.compute_portfolio(conn, "MULTI", date)
        print(f"  invested={res['invested_weight']:.3f} cash={res['cash_residual']:.3f}")
        for _, r in res["funds"].iterrows():
            print(f"    {r['fund_id']:14s} w={r['weight']:.2f} data={r['data_date']} "
                  f"lag={r['lag_days']}d {'STALE' if r['stale'] else ''}")
        print("  dimensions found:", sorted(res["exposures"]))
        for dim in sorted(res["exposures"]):
            df = res["exposures"][dim]
            df = df[df["weight"].abs() > 1e-9]
            cov = res["coverage"].get(dim, 0.0)
            tot = df["weight"].sum()
            # invariant: a dimension's buckets sum to ~the weight of the funds
            # that report it. Source sheets don't always sum to exactly 1
            # (rounding), so allow 1% slack on the coverage.
            assert tot <= cov * 1.01 + 1e-6, (dim, tot, cov)
            top = ", ".join(f"{b}={w:.3f}" for b, w in
                            zip(df["bucket"].head(3), df["weight"].head(3)))
            print(f"    {dim:20s} cov={cov:.2f} sum={tot:.3f} | {top}")
        print("  top holdings:")
        for _, r in res["holdings"].head(3).iterrows():
            print(f"    {r['security'][:32]:32s} {r['weight']:.4f}")
        print(f"  [ok] every dimension sums within its coverage")

    print("\n== Geography (country) evolution (monthly) ==")
    dates = compute.date_steps("2026-01-01", "2026-06-20", 1)
    print("  date steps:", dates)
    evo = compute.compute_evolution(conn, "MULTI", "geography_country", dates)
    print(compute.top_columns(evo, 6).round(4).to_string())
    assert list(evo.index) == dates, (list(evo.index), dates)

    out = BASE / "sample_results.xlsx"
    charts.render_results(compute.compute_portfolio(conn, "MULTI", "2026-06-20"), str(out))
    print(f"\n[ok] wrote charts workbook -> {out}")
    evo_out = BASE / "sample_evolution.xlsx"
    charts.render_evolution(compute.top_columns(evo, 8), str(evo_out),
                            "Geography (country) evolution")
    print(f"[ok] wrote evolution workbook -> {evo_out}")

    conn.close()
    print("\nALL CHECKS PASSED")


if __name__ == "__main__":
    main()
