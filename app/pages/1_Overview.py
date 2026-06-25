"""Overview: national and per-ICB trends over time."""
from __future__ import annotations

import plotly.graph_objects as go
import streamlit as st

import data as D
from ui import NHS_BLUE, NHS_RED, fmt_int, fmt_pct, fmt_weeks

st.set_page_config(page_title="Overview · NHS RTT", page_icon="📈", layout="wide")
st.title("📈 Overview")

# ---- Controls -------------------------------------------------------------
icb_df = D.icbs()
scope_options = ["England (national)"] + icb_df[D.ICB_NAME].tolist()
col_a, col_b = st.columns([2, 1])
scope = col_a.selectbox("Area", scope_options, index=0)
specialty = col_b.selectbox("Specialty", D.specialties(), index=D.specialties().index(D.TOTAL_SPECIALTY))

if scope == "England (national)":
    series = D.national(specialty)
    scope_label = "England"
else:
    code = icb_df.loc[icb_df[D.ICB_NAME] == scope, D.ICB_CODE].iloc[0]
    series = D.icb_series(code, specialty)
    scope_label = scope.replace("NHS ", "").replace(" INTEGRATED CARE BOARD", "").title()

if series.empty:
    st.warning("No data for that combination.")
    st.stop()

latest, first = series.iloc[-1], series.iloc[0]

# ---- KPI row --------------------------------------------------------------
k1, k2, k3, k4 = st.columns(4)
k1.metric("Incomplete pathways", fmt_int(latest[D.INCOMPLETE]),
          delta=fmt_int(latest[D.INCOMPLETE] - first[D.INCOMPLETE]), delta_color="inverse")
k2.metric("Within 18 weeks", fmt_pct(latest[D.PCT18]),
          delta=f"{(latest[D.PCT18]-first[D.PCT18])*100:+.1f} pts")
k3.metric("52+ week waiters", fmt_int(latest[D.W52]),
          delta=fmt_int(latest[D.W52] - first[D.W52]), delta_color="inverse")
k4.metric("Median wait", fmt_weeks(latest[D.MEDIAN]))
st.caption(f"{scope_label} · {specialty} · change shown vs {D.months()[0]}")

# ---- Chart 1: incomplete pathways ----------------------------------------
st.subheader("Incomplete pathways (waiting list size)")
fig = go.Figure()
fig.add_scatter(x=series.index, y=series[D.INCOMPLETE], mode="lines",
                line=dict(color=NHS_BLUE, width=2.5), fill="tozeroy",
                fillcolor="rgba(0,94,184,0.10)", name="Incomplete pathways")
fig.update_layout(height=320, margin=dict(t=10, b=10), yaxis_title="Patients", xaxis_title=None)
st.plotly_chart(fig, use_container_width=True)

# ---- Chart 2: % within 18 weeks vs 92% target ----------------------------
st.subheader("% within 18 weeks vs the 92% standard")
fig2 = go.Figure()
fig2.add_scatter(x=series.index, y=series[D.PCT18] * 100, mode="lines",
                 line=dict(color=NHS_BLUE, width=2.5), name="% within 18 weeks")
fig2.add_hline(y=D.TARGET_PCT18 * 100, line=dict(color=NHS_RED, dash="dash"),
               annotation_text="92% target", annotation_position="top left")
fig2.update_layout(height=320, margin=dict(t=10, b=10),
                   yaxis_title="% within 18 weeks", yaxis_ticksuffix="%", xaxis_title=None)
st.plotly_chart(fig2, use_container_width=True)

# ---- Chart 3: long waiters (52/65/78+) -----------------------------------
st.subheader("Long waiters (52 / 65 / 78+ weeks)")
fig3 = go.Figure()
for col, name, color in [(D.W52, "52+ weeks", "#ED8B00"),
                         (D.W65, "65+ weeks", "#DA291C"),
                         (D.W78, "78+ weeks", "#7C2855")]:
    fig3.add_scatter(x=series.index, y=series[col], mode="lines", name=name,
                     line=dict(color=color, width=2))
fig3.update_layout(height=320, margin=dict(t=10, b=10), yaxis_title="Patients",
                   xaxis_title=None, legend=dict(orientation="h", y=1.1))
st.plotly_chart(fig3, use_container_width=True)

with st.expander("Show data table"):
    cols = [D.PERIOD, D.INCOMPLETE, D.WITHIN18, D.PCT18, D.MEDIAN, D.P92, D.W52, D.W65, D.W78]
    st.dataframe(series[cols].reset_index(drop=True), use_container_width=True, hide_index=True)
