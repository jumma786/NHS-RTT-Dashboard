# What 41 Months of NHS Waiting-List Data Actually Tells Us

### I turned three and a half years of NHS waiting-times data into an interactive dashboard. The story it tells is more nuanced than the headlines.

---

"NHS waiting lists hit 7 million." You've seen the headline. It's true — but it's also the least interesting thing the data says.

I spent a few weeks building a data pipeline and an interactive dashboard on top of NHS England's **Referral-to-Treatment (RTT)** statistics — the official record of how long patients wait between a GP referral and the start of their treatment. The published data covers October 2022 to March 2026: 41 monthly snapshots, broken down by 42 Integrated Care Boards (ICBs) and 23 treatment specialties.

What I found is a system making genuine, dramatic progress on its worst failures — while barely moving on the metric that matters most to the average patient, and quietly leaving your postcode to decide your odds.

Here's what the data shows, and how I built something to explore it.

---

## The headline number hides the real story

The total waiting list — what the NHS calls "incomplete pathways" — was **6.96 million in October 2022**. By March 2026 it was **7.01 million**. Broadly flat (it actually peaked at 7.75M in August 2023 before coming back down).

So if you only read the top-line number, nothing changed. But underneath, two very different things were happening.

### The good news: the longest waits have been all but eliminated

This is the genuine success story, and it's striking:

| Long waiters | Oct 2022 | Mar 2026 | Change |
|---|---:|---:|---:|
| 52+ weeks | 384,015 | 92,609 | **−76%** |
| 65+ weeks | 157,216 | 4,302 | **−97%** |
| 78+ weeks | 46,210 | 1,045 | **−98%** |

Waiting more than a year and a half for treatment used to happen to tens of thousands of people. It now happens to roughly a thousand. That's a deliberate policy target being hit — the NHS set explicit ambitions to eliminate the longest waits first, and the data shows it worked.

### The stubborn news: the 18-week standard is still far away

The NHS's actual *constitutional* standard is that **92% of patients should start treatment within 18 weeks**.

- October 2022: **60.2%** — roughly 6 in 10
- March 2026: **65.3%** — roughly 7 in 10

A 5-point improvement over three and a half years. Progress, but the target is ~9 in 10. At this rate, the gap is measured in years, not months. The median wait improved from 13.9 to 11.3 weeks — better, but still the better part of three months.

So the system got dramatically better at not failing people *catastrophically*, while the *typical* experience improved only modestly.

---

## Your postcode changes your odds by 21 points

This was the finding that surprised me most. The "92% within 18 weeks" standard is national — it's supposed to mean the same thing everywhere. It doesn't.

- **Best ICB:** Gloucestershire — **74.3%** within 18 weeks
- **Worst ICB:** Mid & South Essex — **53.1%**

That's a **21-percentage-point gap** for the same national standard. If you need elective treatment, where you happen to live is one of the biggest factors in how long you'll wait — and that's exactly the kind of inequality that a national average completely hides. Mapping it makes it impossible to ignore, which is why the dashboard includes choropleth maps of every ICB.

## The backlog is concentrated, not spread evenly

Two more things the breakdown reveals:

**By volume**, three specialties carry most of the load — Trauma & Orthopaedics (818k), General Medicine (612k), and Ophthalmology (602k).

**By performance**, the weakest 18-week results are in Oral Surgery (56%), Plastic Surgery (57%), and ENT (58%).

These aren't the same lists, which matters for policy: the biggest *queues* and the worst *waits* are in different places. Throwing resources at the largest specialty wouldn't necessarily fix the worst patient experience.

---

## How I built it

The analysis is only as good as the pipeline underneath it, and NHS open data is messier than it looks.

### The data problem

NHS England publishes RTT data as monthly Excel workbooks — one giant `.xlsx` per month, each with thousands of rows across commissioners and specialties. To analyse trends, you need all 41 months stacked into one tidy dataset. A few wrinkles made that non-trivial:

- **April 2023 is missing** from the modern series — it was only ever released in a legacy `.xls` format. I exclude it from most views and linearly interpolate it inside the forecasting model so the time series stays well-defined.
- **A handful of early months** are missing data for specific trusts (Frimley Health, Manchester University), which has to be handled explicitly rather than silently dropped.
- **National totals don't equal the sum of ICBs.** The all-England figure includes ~127k pathways of NHS England's *direct* specialised-commissioning activity that the ICB breakdown excludes. Getting this wrong means your national numbers are quietly off — a caveat I had to dig into and document.

The ETL pipeline (Python + pandas) reads the raw workbooks, normalises them, and emits three clean CSVs: national totals, ICB-level detail, and a stacked file the app reads directly. The whole thing is reproducible from the raw source with one command.

### The app

The front end is a multi-page **Streamlit** app with five views:

1. **Overview** — national and per-ICB trends: list size, % within 18 weeks vs the 92% line, long waiters, median wait.
2. **Specialties** — which of the 23 treatment functions drive the backlog, by volume and by performance.
3. **Policy** — 52/65/78-week long-waiter tracking against the elimination ambitions, with an ICB league table.
4. **Forecast** — Holt's exponential-smoothing projection per area and metric, with a "when is the 92% standard met?" read-out.
5. **Geography** — ICB choropleth maps (joined to ONS boundary data), best/worst tables, and a best-vs-worst inequality trend.

**Stack:** Python, pandas, NumPy, Streamlit, Plotly, statsmodels (forecasting), scikit-learn. Deployed for free on Streamlit Community Cloud.

---

## What I took away from it

Three lessons, beyond the NHS findings themselves:

1. **Aggregates lie by omission.** "7 million" and "65%" are both true and both hide the real structure — the collapse in long waits, the regional inequality, the specialty concentration. The value of a dashboard isn't the headline; it's letting someone *disaggregate* until the story appears.

2. **The boring 80% is the data engineering.** The interesting charts took an afternoon. Reconciling national vs ICB totals, handling the missing month, and documenting the caveats took most of the time — and that's the part that determines whether the conclusions are trustworthy.

3. **Make the caveats first-class.** Every number in the app is tied to a documented data-quality note. For public-sector data especially, "here's a clean chart" without "here's what's missing and why" is how you mislead people with a straight face.

---

**▶ Try the live dashboard:** https://nhs-rtt-dashboard-zen4ketnck9383sktsgg6h.streamlit.app/

*Data source: NHS England RTT Waiting Times, "Incomplete Commissioner" monthly files, ICB level, Oct 2022 – Mar 2026.*

---

*If you found this useful, I'd love to hear which view surprised you most — and I'm always up for talking data, healthcare analytics, or Python.*
