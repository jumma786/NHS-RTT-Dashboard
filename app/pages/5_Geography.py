"""Geography: ICB choropleth maps and regional inequality over time."""
from __future__ import annotations

import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import data as D
import geo as G
from ui import NHS_BLUE, download_csv, fmt_int, fmt_pct

st.set_page_config(page_title="Geography · NHS RTT", page_icon="🗺️", layout="wide")
st.title("🗺️ Geography & regional inequality")

METRICS = {
    "% within 18 weeks": (D.PCT18, "RdYlGn", "%", 100),
    "Incomplete pathways": (D.INCOMPLETE, "Reds", "", 1),
    "52+ week waiters": (D.W52, "Reds", "", 1),
    "Median wait (weeks)": (D.MEDIAN, "Reds", "", 1),
}

c1, c2 = st.columns([1, 1])
month = c1.select_slider("Month", options=D.months(), value=D.latest_month())
metric_label = c2.selectbox("Metric", list(METRICS))
col, cscale, suffix, scale = METRICS[metric_label]

snap = G.attach(D.snapshot(month, D.TOTAL_SPECIALTY))
snap["value"] = snap[col] * scale

# ---- Choropleth -----------------------------------------------------------
fig = px.choropleth(
    snap, geojson=G.geojson(), locations="geo_name", featureidkey=f"properties.{G.GEO_KEY}",
    color="value", color_continuous_scale=cscale,
    hover_name="ICB Short",
    hover_data={"value": ":,.1f", "geo_name": False},
)
fig.update_geos(fitbounds="locations", visible=False)
fig.update_layout(height=620, margin=dict(t=10, b=10, l=10, r=10),
                  coloraxis_colorbar_title=metric_label.replace(" ", "<br>"))
st.plotly_chart(fig, use_container_width=True)
st.caption(f"All-specialty figures by ICB, {month}. "
           + ("Greener = better (closer to the 92% standard)." if col == D.PCT18
              else "Darker = higher / worse."))

# ---- Best / worst -----------------------------------------------------------
asc = col == D.PCT18  # for %18wk, lowest is worst
ranked = snap.sort_values(col, ascending=asc)
fmt = fmt_pct if col == D.PCT18 else fmt_int
b1, b2 = st.columns(2)
b1.markdown("**Worst 5**")
b1.dataframe(ranked.head(5)[["ICB Short", col]].rename(columns={"ICB Short": "ICB", col: metric_label}),
             hide_index=True, use_container_width=True)
b2.markdown("**Best 5**")
b2.dataframe(ranked.tail(5)[::-1][["ICB Short", col]].rename(columns={"ICB Short": "ICB", col: metric_label}),
             hide_index=True, use_container_width=True)

dl = ranked[["ICB Short", D.INCOMPLETE, D.WITHIN18, D.PCT18, D.MEDIAN, D.W52, D.W65, D.W78]].rename(
    columns={"ICB Short": "ICB"}).reset_index(drop=True)
download_csv(dl, f"rtt_geography_by_icb_{month}")

# ---- Inequality over time (is the gap widening?) --------------------------
st.subheader("Is regional inequality widening or narrowing?")
st.caption("Spread of ICB '% within 18 weeks' each month: the gap between the best "
           "and worst ICBs, and the 10th–90th percentile range.")

df = D.load()
icb_rows = df[(df[D.ICB_NAME] != D.NATIONAL_NAME) & (df[D.TF] == D.TOTAL_SPECIALTY)]
g = icb_rows.groupby("Date")[D.PCT18]
spread = (g.max() - g.min()) * 100
p10, p90 = g.quantile(0.10) * 100, g.quantile(0.90) * 100

fig2 = go.Figure()
fig2.add_scatter(x=p90.index, y=p90.values, line=dict(width=0), showlegend=False, hoverinfo="skip")
fig2.add_scatter(x=p10.index, y=p10.values, fill="tonexty", fillcolor="rgba(0,94,184,0.12)",
                 line=dict(width=0), name="10th–90th percentile")
fig2.add_scatter(x=spread.index, y=spread.values, line=dict(color=NHS_BLUE, width=2.5),
                 name="Max–min gap (pts)")
fig2.update_layout(height=360, margin=dict(t=10), yaxis_title="Percentage points",
                   xaxis_title=None, legend=dict(orientation="h", y=1.12))
st.plotly_chart(fig2, use_container_width=True)

trend = "widening" if spread.iloc[-1] > spread.iloc[0] else "narrowing"
st.info(f"The best-vs-worst ICB gap in % within 18 weeks is **{spread.iloc[-1]:.0f} points** "
        f"in {D.latest_month()} (was {spread.iloc[0]:.0f} points in {D.months()[0]}) — "
        f"overall **{trend}**.")
