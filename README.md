# NHS RTT Waiting Times Dashboard

Interactive Excel dashboard analysing NHS England Referral to Treatment (RTT) waiting times at ICB level.

## Data Source

NHS England RTT waiting times — "Incomplete Commissioner" monthly files, ICB level, October 2022 to March 2026 (42 months).

Source: https://www.england.nhs.uk/statistics/statistical-work-areas/rtt-waiting-times/

### Known Data Quality Caveats

- **Oct 2022 & Nov 2022**: Missing data for Frimley Health NHS FT (RDU) and Manchester University NHS FT (R0A).
- **Dec 2022**: Missing data for Manchester University NHS FT (R0A).

## Folder Structure

| Folder | Contents |
|--------|----------|
| `data/raw/` | 42 monthly source XLSX files (placed manually) |
| `data/processed/` | Cleaned and stacked output |
| `scripts/` | Data cleaning / ETL scripts |
| `dashboard/` | Final Excel dashboard workbook |
| `report/` | Project report drafts (PDF / Word) |
| `notes/` | Data quality notes and known issues |

## Deliverables

- Single `.xlsx` workbook containing the dataset and interactive dashboard.
- Separate written report (`.pdf` or `.docx`).

## Note

This is an Excel-only project per assignment rules (Skill Versed "UK Data Dashboard Project"). Any Python in this repository is used solely for data preparation — the final dashboard and all analysis are built in Excel.
