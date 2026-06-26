"""Policy: long-waiter (52 / 65 / 78+ week) target tracking."""
from __future__ import annotations

import plotly.graph_objects as go
import streamlit as st

import data as D
from ui import NHS_AMBER, NHS_RED, download_csv, fmt_int

st.set_page_config(page_title="Policy · NHS RTT", page_icon="🎯", layout="wide")
st.title("🎯 Long-waiter policy tracking")

st.markdown(
    "NHS England set successive targets to eliminate the longest waits. This page "
    "tracks national progress against the **78**, **65** and **52-week** marks."
)

nat = D.national()  # national Total
latest, first = nat.iloc[-1], nat.iloc[0]

k1, k2, k3 = st.columns(3)
k1.metric("78+ week waiters", fmt_int(latest[D.W78]),
          delta=fmt_int(latest[D.W78] - first[D.W78]), delta_color="inverse")
k2.metric("65+ week waiters", fmt_int(latest[D.W65]),
          delta=fmt_int(latest[D.W65] - first[D.W65]), delta_color="inverse")
k3.metric("52+ week waiters", fmt_int(latest[D.W52]),
          delta=fmt_int(latest[D.W52] - first[D.W52]), delta_color="inverse")
st.caption(f"National · change vs {D.months()[0]}")

# ---- National long-waiter trend with policy milestones -------------------
st.subheader("National long-waiter numbers over time")
fig = go.Figure()
for col, name, color in [(D.W52, "52+ weeks", NHS_AMBER),
                         (D.W65, "65+ weeks", NHS_RED),
                         (D.W78, "78+ weeks", "#7C2855")]:
    fig.add_scatter(x=nat.index, y=nat[col], name=name, mode="lines",
                    line=dict(color=color, width=2.5))

# Policy milestone markers (target dates for elimination ambitions).
milestones = {
    "2023-04-01": "78-wk ambition",
    "2024-03-01": "65-wk ambition",
    "2025-03-01": "52-wk ambition",
}
for date, label in milestones.items():
    fig.add_vline(x=date, line=dict(color="#768692", dash="dot"))
    fig.add_annotation(x=date, y=1, yref="paper", text=label, showarrow=False,
                       font=dict(size=10, color="#768692"), textangle=-90, xshift=-8)
fig.update_layout(height=420, margin=dict(t=20), yaxis_title="Patients",
                  xaxis_title=None, legend=dict(orientation="h", y=1.12))
st.plotly_chart(fig, use_container_width=True)

# ---- ICB league table for latest month -----------------------------------
st.subheader(f"ICB long-waiter league table · {D.latest_month()}")
snap = D.snapshot(D.latest_month())[[D.ICB_NAME, "ICB Short", D.INCOMPLETE, D.W52, D.W65, D.W78, D.PCT18]].copy()
snap["52+ per 10k list"] = (snap[D.W52] / snap[D.INCOMPLETE] * 10000).round(0)
snap = snap.sort_values(D.W52, ascending=False)
show = snap[["ICB Short", D.INCOMPLETE, D.W52, D.W65, D.W78, "52+ per 10k list"]].rename(
    columns={"ICB Short": "ICB", D.INCOMPLETE: "Incomplete", D.W52: "52+", D.W65: "65+", D.W78: "78+"})
st.dataframe(show, use_container_width=True, hide_index=True,
             column_config={c: st.column_config.NumberColumn(format="%d")
                            for c in ["Incomplete", "52+", "65+", "78+", "52+ per 10k list"]})
download_csv(show, f"rtt_policy_league_table_{D.latest_month()}")

n_with_78 = int((snap[D.W78] > 0).sum())
st.caption(f"{n_with_78} of {len(snap)} ICBs still report 78+ week waiters in {D.latest_month()}. "
           "'52+ per 10k list' normalises long waits by list size for fair comparison.")
