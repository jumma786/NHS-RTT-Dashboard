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


def plain_ratio(frac) -> str:
    """Turn a 0–1 fraction into plain English, e.g. 0.653 -> 'about 7 in 10'."""
    try:
        n = round(float(frac) * 10)
        return f"about {n} in 10"
    except (TypeError, ValueError):
        return "–"


# Plain-English definitions of the NHS waiting-list jargon used across the app.
_GLOSSARY = [
    ("Referral to Treatment (RTT)",
     "The wait from being referred (usually by a GP) to starting hospital treatment."),
    ("Incomplete pathway",
     "A person still waiting to start treatment — i.e. someone currently on the waiting list."),
    ("ICB (Integrated Care Board)",
     "The local NHS body that plans and pays for healthcare in your area. There are 42 across England."),
    ("18-week standard",
     "The NHS aims for at least 92% (about 9 in 10) of patients to start treatment within 18 weeks of referral."),
    ("Long waiters (52 / 65 / 78 weeks)",
     "People who have waited more than a year, 15 months, or 18 months — the NHS is working to eliminate these."),
    ("Median wait",
     "The typical wait: half of patients wait less than this, half wait more."),
    ("92nd percentile wait",
     "Close to the longest waits: 92% of patients wait less than this figure."),
]


def glossary(label: str = "ℹ️ New to these terms? Plain-English glossary") -> None:
    """Render a collapsible plain-English glossary of NHS waiting-list terms."""
    import streamlit as st

    with st.expander(label):
        for term, meaning in _GLOSSARY:
            st.markdown(f"**{term}** — {meaning}")


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
