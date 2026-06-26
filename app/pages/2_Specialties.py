"""Specialties: which treatment functions drive the backlog."""
from __future__ import annotations

import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import data as D
from ui import NHS_BLUE, NHS_RED, download_csv, fmt_int, fmt_pct, glossary

st.set_page_config(page_title="Specialties · NHS RTT", page_icon="🩺", layout="wide")
st.title("🩺 Specialties")
st.caption("Which areas of treatment (e.g. hip/knee surgery, eye care) have the most people waiting.")
glossary()

month = st.select_slider("Month", options=D.months(), value=D.latest_month())

brk = D.specialty_breakdown(month, national_only=True).copy()
brk["spec"] = brk[D.TF].str.replace(" Service", "", regex=False).str.replace(" Services", "", regex=False)

c1, c2 = st.columns(2)

# Biggest backlogs by volume
top_vol = brk.sort_values(D.INCOMPLETE, ascending=True)
fig = go.Figure(go.Bar(x=top_vol[D.INCOMPLETE], y=top_vol["spec"], orientation="h",
                       marker_color=NHS_BLUE,
                       text=[fmt_int(v) for v in top_vol[D.INCOMPLETE]], textposition="auto"))
fig.update_layout(height=560, margin=dict(t=10, l=10), title="Incomplete pathways by specialty",
                  xaxis_title="Patients")
c1.plotly_chart(fig, use_container_width=True)

# Worst 18-week performance
worst = brk.sort_values(D.PCT18, ascending=True)
colors = [NHS_RED if v < D.TARGET_PCT18 else NHS_BLUE for v in worst[D.PCT18]]
fig2 = go.Figure(go.Bar(x=worst[D.PCT18] * 100, y=worst["spec"], orientation="h",
                        marker_color=colors,
                        text=[fmt_pct(v) for v in worst[D.PCT18]], textposition="auto"))
fig2.add_vline(x=D.TARGET_PCT18 * 100, line=dict(color=NHS_RED, dash="dash"))
fig2.update_layout(height=560, margin=dict(t=10, l=10),
                   title="% within 18 weeks by specialty (92% target)",
                   xaxis_title="% within 18 weeks", xaxis_ticksuffix="%")
c2.plotly_chart(fig2, use_container_width=True)

st.caption(f"National figures for {month}. Bars in red are below the 92% standard.")
download_csv(brk.drop(columns=["spec"]).reset_index(drop=True), f"rtt_specialties_{month}")

# ---- Trend of the worst specialties over time ----------------------------
st.subheader("Trend: % within 18 weeks for the lowest-performing specialties")
worst_specs = worst[brk[D.TF] != D.TOTAL_SPECIALTY].head(5)[D.TF].tolist()
df = D.load()
trend = df[(df[D.ICB_NAME] == D.NATIONAL_NAME) & (df[D.TF].isin(worst_specs))]
figt = px.line(trend, x="Date", y=trend[D.PCT18] * 100, color=D.TF)
figt.add_hline(y=D.TARGET_PCT18 * 100, line=dict(color=NHS_RED, dash="dash"))
figt.update_layout(height=380, margin=dict(t=10), yaxis_title="% within 18 weeks",
                   yaxis_ticksuffix="%", xaxis_title=None,
                   legend=dict(title=None, orientation="h", y=-0.2))
st.plotly_chart(figt, use_container_width=True)
