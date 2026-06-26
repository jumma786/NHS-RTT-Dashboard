"""Small shared UI helpers: number formatting and a consistent Plotly look."""
from __future__ import annotations

import plotly.io as pio

NHS_BLUE = "#005EB8"
NHS_DARK = "#003087"
NHS_AQUA = "#00A499"
NHS_RED = "#DA291C"
NHS_AMBER = "#ED8B00"
NHS_GREY = "#768692"

# A clean default template for every chart in the app.
pio.templates["nhs"] = pio.templates["plotly_white"]
pio.templates["nhs"].layout.colorway = [NHS_BLUE, NHS_AQUA, NHS_AMBER, NHS_RED, NHS_GREY, NHS_DARK]
pio.templates["nhs"].layout.font.family = "Arial, Helvetica, sans-serif"
pio.templates.default = "nhs"


def fmt_int(x) -> str:
    try:
        return f"{int(round(float(x))):,}"
    except (TypeError, ValueError):
        return "–"


def fmt_pct(x, dp: int = 1) -> str:
    """Format a 0–1 fraction as a percentage string."""
    try:
        return f"{float(x) * 100:.{dp}f}%"
    except (TypeError, ValueError):
        return "–"


def fmt_weeks(x, dp: int = 1) -> str:
    try:
        return f"{float(x):.{dp}f} wks"
    except (TypeError, ValueError):
        return "–"


def delta_int(curr, prev) -> str:
    try:
        return f"{int(round(float(curr) - float(prev))):+,}"
    except (TypeError, ValueError):
        return ""


def download_csv(df, filename: str, label: str = "⬇ Download CSV", key=None) -> None:
    """Render a Streamlit button that downloads ``df`` as a CSV file.

    ``filename`` is sanitised so spaces and slashes don't break the download.
    Imported lazily so this module stays usable outside Streamlit.
    """
    import re

    import streamlit as st

    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", filename).strip("_") or "data.csv"
    if not safe.lower().endswith(".csv"):
        safe += ".csv"
    st.download_button(
        label,
        data=df.to_csv(index=False).encode("utf-8"),
        file_name=safe,
        mime="text/csv",
        key=key,
    )
