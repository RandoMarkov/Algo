"""Central configuration: file paths, the three fund-type layout profiles, and
the front-end workbook layout.

Three kinds of fund file exist (Equity / Fixed Income / Alternatives). They
share a `Synthese` identity sheet and an `Inventaire` holdings sheet, but each
reports a different set of exposure sheets. Each profile below lists, per type,
which sheets to read and how. The DB's generic fund_exposures(dimension,
bucket, weight) table absorbs whatever dimensions each type reports.

Adapt to layout changes here, not in ingest.py.
"""
from pathlib import Path

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "funds.db"

# ---------------------------------------------------------------------------
# Synthese sheet (identity + metrics) — common to all three types.
# Labels live in column A, values in column B. Matching is case-insensitive.
# ---------------------------------------------------------------------------
SHEET_SYNTHESIS = "Synthese"
HOLDINGS_SHEET = "Inventaire"

# Synthese labels that are IDENTITY (not metrics). Lowercased keys -> our field.
SYN_IDENTITY = {
    "nom du fonds": "name",
    "isin": "fund_id",        # ISIN is the unique fund id
    "categorie": "category",  # drives type detection; also stored as attribute
    "date": "as_of",
}
# Numeric Synthese values that are NOT meaningful to weight-average across funds
# (levels/counts rather than rates). Excluded from the portfolio metric rollup.
NON_AVERAGEABLE_METRICS = {"aum du fonds", "nombre de titres"}

# Holdings (Inventaire) column headers -> our fields.
HOLD_COLS = {"isin": "ISIN", "security": "Nom du titre", "weight": "Poids"}

# ---------------------------------------------------------------------------
# Exposure-sheet reader spec. One dict per sheet. Defaults:
#   weight_col=2, label_cols=[1], ffill_cols=[], start_row=2
# Rows whose joined label is in SKIP_LABELS (case-insensitive) or whose weight
# cell is non-numeric/blank are skipped. Several sheets may share one dimension
# (e.g. DM + EM geography -> one country-level dimension).
# ---------------------------------------------------------------------------
SKIP_LABELS = {"total"}

def _ex(sheet, dimension, weight_col=2, label_cols=(1,), ffill_cols=(), start_row=2):
    return {"sheet": sheet, "dimension": dimension, "weight_col": weight_col,
            "label_cols": list(label_cols), "ffill_cols": list(ffill_cols),
            "start_row": start_row}

PROFILES = {
    "Equity": {
        "match": ("equity",),
        "exposures": [
            _ex("GeographieRegion", "geography_region"),
            _ex("GeographieDM", "geography_country"),
            _ex("GeographieEM", "geography_country"),
            _ex("DeviseDM", "currency"),
            _ex("DeviseEM", "currency"),
            _ex("Secteur", "sector"),
            _ex("Market Cap Europe", "market_cap_europe"),
            _ex("Market Cap US", "market_cap_us"),
        ],
    },
    "Fixed Income": {
        "match": ("fixed income", "bond", "obligation"),
        "exposures": [
            _ex("Geographie", "geography_country"),
            _ex("Rating", "credit_rating"),
            _ex("Maturite", "maturity"),
            # hierarchical: category in A (filled down), sub in B, weight in C
            _ex("Secteur", "sector_fi", weight_col=3, label_cols=(1, 2), ffill_cols=(1,)),
        ],
    },
    "Alternatives": {
        "match": ("alternative",),
        "exposures": [
            _ex("GeographieDM", "geography_country"),
            _ex("GeographieEM", "geography_country"),
            _ex("Devise", "currency"),
            _ex("Secteur", "sector"),
            _ex("Capitalisation", "market_cap"),  # uses 'Net' column (B)
        ],
    },
}

# Human-readable labels for dimensions (used in Results/Charts).
DIM_LABELS = {
    "geography_region": "Geography (region)",
    "geography_country": "Geography (country)",
    "currency": "Currency",
    "sector": "Sector (GICS)",
    "sector_fi": "Sector (fixed income)",
    "credit_rating": "Credit rating",
    "maturity": "Maturity",
    "market_cap": "Market cap",
    "market_cap_europe": "Market cap (Europe)",
    "market_cap_us": "Market cap (US)",
}


def profile_for_category(category: str) -> str:
    """Map a Synthese 'Categorie' value to a profile name."""
    c = (category or "").lower()
    for name, prof in PROFILES.items():
        if any(token in c for token in prof["match"]):
            return name
    raise ValueError(f"No fund-type profile matches Categorie={category!r}")


# ---------------------------------------------------------------------------
# Front-end workbook layout (used by app.py / xlwings).
# ---------------------------------------------------------------------------
CTRL_SHEET = "Control"
CTRL_CELLS = {
    "fund_file": "B2",
    "portfolio_name": "B3",
    "as_of_date": "B4",
    # evolution inputs (button 4)
    "evo_dimension": "B5",
    "evo_start": "B6",
    "evo_end": "B7",
    "evo_step_months": "B8",
}
STATUS_CELL = "B10"
PORTFOLIO_SHEET = "Portfolio"
RESULTS_SHEET = "Results"
CHARTS_SHEET = "Charts"
EVOLUTION_SHEET = "Evolution"

STALENESS_DAYS = 45
