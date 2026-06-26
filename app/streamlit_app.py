"""NHS RTT Waiting Times — analytics app (entry point).

Run with:  streamlit run app/streamlit_app.py
"""
from __future__ import annotations

import streamlit as st

import data as D
from ui import fmt_int, fmt_pct, fmt_weeks, glossary, plain_ratio

st.set_page_config(
    page_title="NHS RTT Analytics",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🏥 NHS RTT Waiting Times — Analytics")
st.caption(
    "How long people wait to start hospital treatment in England, by local NHS "
    f"area, {D.months()[0]} to {D.latest_month()} ({len(D.months())} months). "
    "\"Incomplete pathways\" = people currently on the waiting list."
)
glossary()

nat = D.national()  # national 'Total' specialty series
latest = nat.iloc[-1]
prev = nat.iloc[-2]

c1, c2, c3, c4 = st.columns(4)
c1.metric(
    "Incomplete pathways",
    fmt_int(latest[D.INCOMPLETE]),
    delta=fmt_int(latest[D.INCOMPLETE] - prev[D.INCOMPLETE]),
    delta_color="inverse",
)
c2.metric(
    "Within 18 weeks",
    fmt_pct(latest[D.PCT18]),
    delta=f"{(latest[D.PCT18]-prev[D.PCT18])*100:+.1f} pts",
)
c3.metric(
    "52+ week waiters",
    fmt_int(latest[D.W52]),
    delta=fmt_int(latest[D.W52] - prev[D.W52]),
    delta_color="inverse",
)
c4.metric("Median wait", fmt_weeks(latest[D.MEDIAN]))

gap = (D.TARGET_PCT18 - latest[D.PCT18]) * 100
st.info(
    f"**In plain terms:** {plain_ratio(latest[D.PCT18])} patients start treatment "
    f"within 18 weeks ({fmt_pct(latest[D.PCT18])}). The NHS aims for **about 9 in 10** "
    f"(the 92% standard) — so we're **{gap:.1f} percentage points** short."
)

st.markdown(
    """
### Explore
Use the sidebar pages:

- **Overview** — national & ICB trends over time
- **Geography** — ICB maps and regional inequality
- **Specialties** — which treatment functions drive the backlog
- **Policy** — 52 / 65 / 78-week long-waiter target tracking
- **Forecast** — projected performance and when the 92% standard might be met

---
*Source: NHS England RTT "Incomplete Commissioner" monthly files. April 2023 is
absent from the published series and is therefore excluded throughout.*
"""
)
