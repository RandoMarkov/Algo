"""SQLite schema, connection, and the 'as-of' query helpers.

The whole application is built around one idea: funds publish data at discrete
dates (often with a lag), so any query for date D must use the *latest data at
or before D* for each fund independently. `latest_date_for` encapsulates that.
"""
import sqlite3
from pathlib import Path

from config import DB_PATH

SCHEMA = """
CREATE TABLE IF NOT EXISTS funds (
    fund_id   TEXT PRIMARY KEY,
    name      TEXT,
    currency  TEXT
);

-- Look-through holdings: each row is one security's weight in a fund snapshot.
CREATE TABLE IF NOT EXISTS fund_holdings (
    fund_id    TEXT,
    as_of_date TEXT,              -- ISO 'YYYY-MM-DD'
    security   TEXT,
    isin       TEXT,
    weight     REAL,              -- fraction of fund NAV (0..1)
    PRIMARY KEY (fund_id, as_of_date, security)
);

-- Generic exposure table: one row per (fund, date, dimension, bucket).
-- dimension is e.g. 'geography' or 'currency'; bucket is 'US', 'EUR', ...
CREATE TABLE IF NOT EXISTS fund_exposures (
    fund_id    TEXT,
    as_of_date TEXT,
    dimension  TEXT,
    bucket     TEXT,
    weight     REAL,
    PRIMARY KEY (fund_id, as_of_date, dimension, bucket)
);

-- Headline NUMERIC metrics from the synthesis sheet (maturity, duration, ...).
CREATE TABLE IF NOT EXISTS fund_metrics (
    fund_id    TEXT,
    as_of_date TEXT,
    metric     TEXT,
    value      REAL,
    PRIMARY KEY (fund_id, as_of_date, metric)
);

-- TEXT attributes from the synthesis sheet (category, benchmark, credit rating).
CREATE TABLE IF NOT EXISTS fund_attributes (
    fund_id    TEXT,
    as_of_date TEXT,
    attribute  TEXT,
    value_text TEXT,
    PRIMARY KEY (fund_id, as_of_date, attribute)
);

CREATE TABLE IF NOT EXISTS portfolios (
    portfolio_id  TEXT PRIMARY KEY,
    name          TEXT,
    base_currency TEXT
);

-- Time-varying allocation. weight applies from effective_date until the next
-- effective_date for the same (portfolio, fund). weight 0 == fund removed.
CREATE TABLE IF NOT EXISTS portfolio_allocations (
    portfolio_id   TEXT,
    fund_id        TEXT,
    effective_date TEXT,
    weight         REAL,
    PRIMARY KEY (portfolio_id, fund_id, effective_date)
);
"""


def connect(db_path: Path | str = DB_PATH) -> sqlite3.Connection:
    """Open (creating if needed) the SQLite DB and ensure the schema exists."""
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA)
    return conn


def latest_date_for(conn, table: str, fund_id: str, target_date: str,
                    dimension: str | None = None) -> str | None:
    """Return the most recent as_of_date <= target_date for a fund in `table`.

    This is the heart of the 'as-of' semantics. Returns None if the fund has no
    data at or before the target date.
    """
    sql = f"SELECT MAX(as_of_date) AS d FROM {table} WHERE fund_id=? AND as_of_date<=?"
    params = [fund_id, target_date]
    if dimension is not None:
        sql += " AND dimension=?"
        params.append(dimension)
    row = conn.execute(sql, params).fetchone()
    return row["d"] if row and row["d"] else None


# --- idempotent writes ------------------------------------------------------
# Re-recording a fund for a date it already has should REPLACE that snapshot,
# not duplicate it. Each loader deletes the (fund, date) slice first.

def replace_snapshot(conn, table: str, fund_id: str, as_of_date: str,
                     rows: list[dict]):
    """Delete the existing (fund_id, as_of_date) rows in `table`, insert `rows`."""
    if not rows:
        return
    conn.execute(f"DELETE FROM {table} WHERE fund_id=? AND as_of_date=?",
                 (fund_id, as_of_date))
    cols = list(rows[0].keys())
    placeholders = ",".join("?" * len(cols))
    conn.executemany(
        f"INSERT INTO {table} ({','.join(cols)}) VALUES ({placeholders})",
        [tuple(r[c] for c in cols) for r in rows],
    )


def upsert_fund(conn, fund_id: str, name: str, currency: str):
    conn.execute(
        """INSERT INTO funds (fund_id, name, currency) VALUES (?,?,?)
           ON CONFLICT(fund_id) DO UPDATE SET name=excluded.name,
                                              currency=excluded.currency""",
        (fund_id, name, currency),
    )


def replace_allocations(conn, portfolio_id: str, name: str, base_currency: str,
                        allocations: list[dict]):
    """Define/replace a portfolio's full allocation table.

    allocations: list of {fund_id, effective_date, weight}.
    """
    conn.execute(
        """INSERT INTO portfolios (portfolio_id, name, base_currency) VALUES (?,?,?)
           ON CONFLICT(portfolio_id) DO UPDATE SET name=excluded.name,
                                                   base_currency=excluded.base_currency""",
        (portfolio_id, name, base_currency),
    )
    conn.execute("DELETE FROM portfolio_allocations WHERE portfolio_id=?",
                 (portfolio_id,))
    conn.executemany(
        """INSERT INTO portfolio_allocations
           (portfolio_id, fund_id, effective_date, weight) VALUES (?,?,?,?)""",
        [(portfolio_id, a["fund_id"], a["effective_date"], a["weight"])
         for a in allocations],
    )
