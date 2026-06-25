"""Shared data layer for the NHS RTT analytics app.

Loads the processed RTT CSVs and exposes tidy helper functions used by every
page. Caching is wired through Streamlit when available but the module also
works standalone (e.g. in notebooks / tests).
"""
from __future__ import annotations

import os
from functools import lru_cache

import pandas as pd

# ---------------------------------------------------------------------------
# Column constants (single source of truth for the processed schema)
# ---------------------------------------------------------------------------
PERIOD = "Period"
ICB_CODE = "ICB Code"
ICB_NAME = "ICB Name"
TFC = "Treatment Function Code"
TF = "Treatment Function"

INCOMPLETE = "Total number of incomplete pathways"
WITHIN18 = "Total within 18 weeks"
PCT18 = "% within 18 weeks"
MEDIAN = "Average (median) waiting time (in weeks)"
P92 = "92nd percentile waiting time (in weeks)"
W52 = "Total 52 plus weeks"
W65 = "Total 65 plus weeks"
W78 = "Total 78 plus weeks"
PCT52 = "% 52 plus weeks"

TOTAL_SPECIALTY = "Total"
NATIONAL_NAME = "NHS ENGLAND"
NATIONAL_CODE = "-"

# 18-week operational standard: 92% of incomplete pathways within 18 weeks.
TARGET_PCT18 = 0.92

METRICS = [INCOMPLETE, WITHIN18, PCT18, MEDIAN, P92, W52, W65, W78, PCT52]

_DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "processed")
_STACKED = os.path.join(_DATA_DIR, "rtt_icb_stacked.csv")


# ---------------------------------------------------------------------------
# Optional Streamlit caching: use st.cache_data if running inside Streamlit,
# otherwise fall back to a plain lru_cache so the module stays importable.
# ---------------------------------------------------------------------------
def _cache(func):
    try:
        import streamlit as st
        from streamlit.runtime import exists as _rt_exists

        if _rt_exists():
            return st.cache_data(show_spinner=False)(func)
    except Exception:
        pass
    return lru_cache(maxsize=None)(func)


@_cache
def load() -> pd.DataFrame:
    """Load the stacked ICB file (ICB rows + NHS ENGLAND national rows).

    Adds a ``Date`` column (month-start Timestamp) for time-series work and
    a short ``ICB Short`` label for charts/maps.
    """
    df = pd.read_csv(_STACKED, dtype={PERIOD: str})
    df["Date"] = pd.to_datetime(df[PERIOD] + "-01", format="%Y-%m-%d")
    df["ICB Short"] = (
        df[ICB_NAME]
        .str.replace("NHS ", "", regex=False)
        .str.replace(" INTEGRATED CARE BOARD", "", regex=False)
        .str.title()
    )
    return df


# ---------------------------------------------------------------------------
# Dimension helpers
# ---------------------------------------------------------------------------
def months() -> list[str]:
    return sorted(load()[PERIOD].unique())


def latest_month() -> str:
    return months()[-1]


def specialties(include_total: bool = True) -> list[str]:
    s = sorted(load()[TF].unique())
    if not include_total:
        s = [x for x in s if x != TOTAL_SPECIALTY]
    return s


def icbs() -> pd.DataFrame:
    """Return distinct ICBs (excludes the national NHS ENGLAND aggregate)."""
    df = load()
    out = (
        df[df[ICB_NAME] != NATIONAL_NAME][[ICB_CODE, ICB_NAME, "ICB Short"]]
        .drop_duplicates()
        .sort_values(ICB_NAME)
        .reset_index(drop=True)
    )
    return out


# ---------------------------------------------------------------------------
# Slicing helpers
# ---------------------------------------------------------------------------
def national(specialty: str = TOTAL_SPECIALTY) -> pd.DataFrame:
    """National (NHS ENGLAND) monthly series for one specialty, Date-indexed."""
    df = load()
    out = df[(df[ICB_NAME] == NATIONAL_NAME) & (df[TF] == specialty)]
    return out.sort_values("Date").set_index("Date")


def icb_series(icb_code: str, specialty: str = TOTAL_SPECIALTY) -> pd.DataFrame:
    """Monthly series for one ICB + specialty, Date-indexed."""
    df = load()
    out = df[(df[ICB_CODE] == icb_code) & (df[TF] == specialty)]
    return out.sort_values("Date").set_index("Date")


def snapshot(month: str, specialty: str = TOTAL_SPECIALTY, national_row: bool = False) -> pd.DataFrame:
    """All ICBs for a given month + specialty (one row per ICB)."""
    df = load()
    mask = (df[PERIOD] == month) & (df[TF] == specialty)
    if not national_row:
        mask &= df[ICB_NAME] != NATIONAL_NAME
    return df[mask].reset_index(drop=True)


def specialty_breakdown(month: str, national_only: bool = True) -> pd.DataFrame:
    """Per-specialty metrics for a month (national by default), excludes 'Total'."""
    df = load()
    mask = (df[PERIOD] == month) & (df[TF] != TOTAL_SPECIALTY)
    mask &= df[ICB_NAME] == NATIONAL_NAME if national_only else df[ICB_NAME] != NATIONAL_NAME
    return df[mask].reset_index(drop=True)


if __name__ == "__main__":
    d = load()
    print("rows:", len(d), "| months:", len(months()), months()[0], "->", latest_month())
    print("ICBs:", len(icbs()), "| specialties:", len(specialties()))
    n = national()
    print("national Total latest %18wk:", round(float(n[PCT18].iloc[-1]) * 100, 1), "%")
