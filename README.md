# NHS RTT Waiting Times — Analytics

Data pipeline and interactive **Streamlit** app analysing NHS England Referral to
Treatment (RTT) waiting times at Integrated Care Board (ICB) level.

The app turns 41 months of national waiting-list data into five analytical views:
national trends, specialty breakdowns, long-waiter policy tracking, forecasting,
and geographic inequality maps.

## Data Source

NHS England RTT waiting times — "Incomplete Commissioner" monthly files, ICB level,
**October 2022 to March 2026** (41 months published).

Source: https://www.england.nhs.uk/statistics/statistical-work-areas/rtt-waiting-times/

ICB boundaries for the maps are the ONS April-2023 super-generalised boundaries
(Open Geography Portal), joined to the data by ICB name (all 42 ICBs match).

### Known Data Quality Caveats

- **April 2023** is absent from the published series (released only in legacy `.xls`
  format) and is therefore excluded throughout. The Forecast page linearly
  interpolates this single interior month internally so the time-series model stays
  well-defined; all other views simply skip it.
- **Oct 2022 & Nov 2022**: missing data for Frimley Health NHS FT (RDU) and
  Manchester University NHS FT (R0A).
- **Dec 2022**: missing data for Manchester University NHS FT (R0A).

## Quick Start

```bash
pip install -r requirements.txt
streamlit run app/streamlit_app.py
```

Then open http://localhost:8501.

> The app reads `data/processed/*.csv`. These are large and git-ignored — regenerate
> them from the raw monthly files with `python scripts/etl.py` (see below).

## The App

| Page | Contents |
|------|----------|
| **Overview** | National & per-ICB trends: waiting list size, % within 18 weeks vs the 92% standard, long waiters, median wait. Filter by area and specialty. |
| **Specialties** | Which of the 23 treatment functions drive the backlog — by volume and by 18-week performance — plus trends for the worst performers. |
| **Policy** | 52 / 65 / 78-week long-waiter tracking against NHS elimination ambitions, with an ICB league table. |
| **Forecast** | Holt's exponential-smoothing projection per area/metric, with an indicative band and a "when is the 92% standard met?" read-out. |
| **Geography** | ICB choropleth maps by metric and month, best/worst tables, and a best-vs-worst inequality trend. |

The earlier Excel workbook (`dashboard/NHS_RTT_Dashboard.xlsx`) and written report
(`report/NHS_RTT_Report.docx`) remain in the repo as standalone deliverables.

## Folder Structure

| Path | Contents |
|------|----------|
| `app/` | Streamlit app — `streamlit_app.py` (entry), `pages/`, `data.py`, `ui.py`, `geo.py` |
| `scripts/` | Data preparation — `etl.py`, plus `build_dashboard.py` / `build_report.py` for the Excel/Word outputs |
| `data/raw/` | Monthly source XLSX files (git-ignored; placed manually / downloaded) |
| `data/processed/` | Cleaned, stacked CSV output (git-ignored; regenerate via `etl.py`) |
| `data/geo/` | ONS ICB boundary GeoJSON for the maps |
| `notebook/` | `RTT Performance Analysis.ipynb` exploratory analysis |
| `dashboard/` | Excel dashboard workbook |
| `report/` | Written report draft |
| `notes/` | Data quality notes |

## Data Pipeline

```bash
python scripts/etl.py            # raw XLSX -> data/processed/*.csv
```

Produces three tidy files:

- `rtt_national.csv` — England totals by treatment function and month
- `rtt_icb.csv` — 42 ICBs × 24 treatment functions × month
- `rtt_icb_stacked.csv` — as above plus an `NHS ENGLAND` national row set (the file
  the app reads)

## Requirements

Python 3.10+. Key packages (see `requirements.txt`): `pandas`, `numpy`, `streamlit`,
`plotly`, `statsmodels`, `scikit-learn`, `geopandas`, `openpyxl`.
