"""Forecast: project a metric forward and flag when the 92% standard is met."""
from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from statsmodels.tsa.holtwinters import ExponentialSmoothing

import data as D
from ui import NHS_BLUE, NHS_RED

st.set_page_config(page_title="Forecast · NHS RTT", page_icon="🔮", layout="wide")
st.title("🔮 Forecast")
st.caption("Indicative projection using Holt's exponential smoothing (trend). "
           "Not an official NHS forecast — for exploratory analysis only.")

METRIC_CHOICES = {
    "% within 18 weeks": D.PCT18,
    "Incomplete pathways": D.INCOMPLETE,
    "52+ week waiters": D.W52,
}

c1, c2, c3 = st.columns(3)
icb_df = D.icbs()
scope = c1.selectbox("Area", ["England (national)"] + icb_df[D.ICB_NAME].tolist())
metric_label = c2.selectbox("Metric", list(METRIC_CHOICES))
horizon = c3.slider("Forecast horizon (months)", 6, 36, 18)
damped = st.checkbox("Damped trend (more conservative)", value=True)

metric = METRIC_CHOICES[metric_label]
if scope == "England (national)":
    series = D.national()
else:
    code = icb_df.loc[icb_df[D.ICB_NAME] == scope, D.ICB_CODE].iloc[0]
    series = D.icb_series(code)

y = series[metric].astype(float)
y.index = pd.DatetimeIndex(series.index).to_period("M").to_timestamp()
# Reindex to a gap-free monthly axis; April 2023 is absent from the NHS series,
# so interpolate that single interior month to keep the model well-defined.
y = y.asfreq("MS").interpolate(method="linear").ffill().bfill()

# ---- Fit Holt's linear (optionally damped) trend model -------------------
model = ExponentialSmoothing(y, trend="add", damped_trend=damped,
                             initialization_method="estimated").fit()
fc_index = pd.date_range(y.index[-1] + pd.offsets.MonthBegin(1), periods=horizon, freq="MS")
fc = pd.Series(np.asarray(model.forecast(horizon)), index=fc_index)

# Residual-based indicative interval (widening with horizon).
resid_std = float(np.std(model.resid))
steps = np.arange(1, horizon + 1)
band = 1.28 * resid_std * np.sqrt(steps)  # ~80% indicative band
upper, lower = fc.values + band, fc.values - band

is_pct = metric == D.PCT18
scale = 100 if is_pct else 1
suffix = "%" if is_pct else ""

fig = go.Figure()
fig.add_scatter(x=y.index, y=y.values * scale, mode="lines",
                line=dict(color=NHS_BLUE, width=2.5), name="Actual")
fig.add_scatter(x=fc.index, y=fc.values * scale, mode="lines",
                line=dict(color=NHS_BLUE, width=2.5, dash="dash"), name="Forecast")
fig.add_scatter(x=list(fc.index) + list(fc.index[::-1]),
                y=list(upper * scale) + list(lower[::-1] * scale),
                fill="toself", fillcolor="rgba(0,94,184,0.12)",
                line=dict(color="rgba(0,0,0,0)"), name="~80% band", hoverinfo="skip")
if is_pct:
    fig.add_hline(y=D.TARGET_PCT18 * 100, line=dict(color=NHS_RED, dash="dot"),
                  annotation_text="92% target")
fig.update_layout(height=440, margin=dict(t=20), yaxis_title=metric_label,
                  yaxis_ticksuffix=suffix, xaxis_title=None,
                  legend=dict(orientation="h", y=1.1))
st.plotly_chart(fig, use_container_width=True)

# ---- Headline read-out ----------------------------------------------------
end_val = fc.iloc[-1]
end_txt = f"{end_val*100:.1f}%" if is_pct else f"{end_val:,.0f}"
st.metric(f"Projected {metric_label} in {fc.index[-1]:%b %Y}", end_txt)

if is_pct:
    hits = fc[fc >= D.TARGET_PCT18]
    if len(hits):
        st.success(f"On this trajectory the **92% standard is reached around "
                   f"{hits.index[0]:%B %Y}**.")
    else:
        st.warning(f"On this trajectory the 92% standard is **not reached within "
                   f"{horizon} months** (projected {end_txt} by {fc.index[-1]:%b %Y}).")

with st.expander("Model details"):
    st.write({"model": "ExponentialSmoothing (additive trend)",
              "damped_trend": damped, "AIC": round(float(model.aic), 1),
              "residual_std": round(resid_std, 4), "n_obs": int(len(y))})
