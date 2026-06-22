"""Portfolio look-through: holdings, exposures (any dimension) and metrics
as-of a date, for a time-varying basket of funds of mixed types.

All 'as-of' bookkeeping (latest snapshot <= target, per fund) lives in
db.latest_date_for. Exposure dimensions are discovered from the data, so funds
of different types (each reporting different dimensions) combine cleanly.
"""
from calendar import monthrange
from datetime import date, datetime

import pandas as pd

import config
import db


def date_steps(start: str, end: str, step_months: int = 1) -> list[str]:
    """ISO dates from start to end inclusive, stepping `step_months` months.

    Keeps the start day-of-month (clamped to month length) and always includes
    the end date so the latest point is shown.
    """
    step_months = max(1, int(step_months))
    s = datetime.strptime(start, "%Y-%m-%d").date()
    e = datetime.strptime(end, "%Y-%m-%d").date()
    out, y, m, d = [], s.year, s.month, s.day
    cur = s
    while cur <= e:
        out.append(cur.isoformat())
        m0 = m - 1 + step_months
        y, m = y + m0 // 12, m0 % 12 + 1
        cur = date(y, m, min(d, monthrange(y, m)[1]))
    if not out or out[-1] != e.isoformat():
        out.append(e.isoformat())
    return out


def top_columns(df: pd.DataFrame, n: int = 8) -> pd.DataFrame:
    """Keep the n largest-weight columns; collapse the rest into 'Other'."""
    if df.shape[1] <= n:
        return df
    order = df.max().sort_values(ascending=False).index
    keep, rest = list(order[:n]), list(order[n:])
    out = df[keep].copy()
    out["Other"] = df[rest].sum(axis=1)
    return out


def _allocation_asof(conn, portfolio_id, target_date) -> dict:
    """{fund_id: weight} using the latest effective_date <= target per fund."""
    rows = conn.execute(
        """
        WITH ranked AS (
            SELECT fund_id, weight, effective_date,
                   ROW_NUMBER() OVER (PARTITION BY fund_id
                                      ORDER BY effective_date DESC) AS rn
            FROM portfolio_allocations
            WHERE portfolio_id=? AND effective_date<=?
        )
        SELECT fund_id, weight FROM ranked WHERE rn=1 AND weight<>0
        """,
        (portfolio_id, target_date),
    ).fetchall()
    return {r["fund_id"]: r["weight"] for r in rows}


def _days_between(d1, d2) -> int:
    return abs((datetime.strptime(d1, "%Y-%m-%d")
                - datetime.strptime(d2, "%Y-%m-%d")).days)


def _df(d: dict, key_name: str) -> pd.DataFrame:
    if not d:
        return pd.DataFrame(columns=[key_name, "weight"])
    return pd.DataFrame(sorted(d.items(), key=lambda kv: -kv[1]),
                        columns=[key_name, "weight"])


def compute_portfolio(conn, portfolio_id, target_date) -> dict:
    """Consolidated look-through for one portfolio on one date.

    Returns:
      holdings   : DataFrame(security, weight)         look-through
      exposures  : {dimension: DataFrame(bucket, weight)}   all dimensions found
      metrics    : DataFrame(metric, value)            weighted average
      funds      : DataFrame(fund_id, weight, data_date, lag_days, stale)
      coverage   : {dimension: invested-weight covered by funds reporting it}
    plus scalars invested_weight, cash_residual.
    """
    weights = _allocation_asof(conn, portfolio_id, target_date)
    invested = sum(weights.values())

    holdings, metrics = {}, {}
    exposures = {}                      # dimension -> {bucket: weight}
    coverage = {}                       # dimension -> summed fund weight reporting it
    fund_rows = []

    for fund_id, w in weights.items():
        # --- holdings (look-through) ---
        hd = db.latest_date_for(conn, "fund_holdings", fund_id, target_date)
        if hd:
            for r in conn.execute(
                "SELECT security, weight FROM fund_holdings WHERE fund_id=? AND as_of_date=?",
                (fund_id, hd)):
                holdings[r["security"]] = holdings.get(r["security"], 0.0) + w * r["weight"]

        # --- exposures: one snapshot date per fund, all dimensions from it ---
        ed = db.latest_date_for(conn, "fund_exposures", fund_id, target_date)
        if ed:
            dims_here = set()
            for r in conn.execute(
                """SELECT dimension, bucket, weight FROM fund_exposures
                   WHERE fund_id=? AND as_of_date=?""", (fund_id, ed)):
                dim, bucket = r["dimension"], r["bucket"]
                exposures.setdefault(dim, {})[bucket] = (
                    exposures.setdefault(dim, {}).get(bucket, 0.0) + w * r["weight"])
                dims_here.add(dim)
            for dim in dims_here:
                coverage[dim] = coverage.get(dim, 0.0) + w

        # --- metrics (weighted, averageable ones only) ---
        md = db.latest_date_for(conn, "fund_metrics", fund_id, target_date)
        if md:
            for r in conn.execute(
                "SELECT metric, value FROM fund_metrics WHERE fund_id=? AND as_of_date=?",
                (fund_id, md)):
                if r["metric"].lower() in config.NON_AVERAGEABLE_METRICS:
                    continue
                metrics[r["metric"]] = metrics.get(r["metric"], 0.0) + w * r["value"]

        used = ed or hd or md
        fund_rows.append({
            "fund_id": fund_id, "weight": w, "data_date": used,
            "lag_days": _days_between(used, target_date) if used else None,
            "stale": used is not None and _days_between(used, target_date) > config.STALENESS_DAYS,
        })

    # weighted-average metrics: divide weighted sums by invested weight
    if invested:
        metrics = {k: v / invested for k, v in metrics.items()}

    return {
        "portfolio_id": portfolio_id,
        "as_of_date": target_date,
        "invested_weight": invested,
        "cash_residual": 1.0 - invested,
        "holdings": _df(holdings, "security"),
        "exposures": {dim: _df(b, "bucket") for dim, b in exposures.items()},
        "coverage": coverage,
        "metrics": (pd.DataFrame(sorted(metrics.items()), columns=["metric", "value"])
                    if metrics else pd.DataFrame(columns=["metric", "value"])),
        "funds": pd.DataFrame(fund_rows),
    }


def compute_evolution(conn, portfolio_id, dimension, dates) -> pd.DataFrame:
    """Evolution of one exposure dimension over time: rows=date, cols=buckets."""
    series = {}
    for d in dates:
        res = compute_portfolio(conn, portfolio_id, d)
        df = res["exposures"].get(dimension, pd.DataFrame(columns=["bucket", "weight"]))
        series[d] = dict(zip(df["bucket"], df["weight"]))
    out = pd.DataFrame(series).T.fillna(0.0)
    out.index.name = "date"
    return out
