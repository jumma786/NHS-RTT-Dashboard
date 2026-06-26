"""
ETL script: reads monthly NHS RTT 'Incomplete Commissioner' XLSX files.

Extracts two sheets from each file:
  - National  → national-level totals by treatment function
  - ICB       → ICB-level data (includes an NHS England aggregate row)

Outputs to data/processed/:
  - rtt_national.csv      National-level rows
  - rtt_icb_stacked.csv   All ICB-sheet rows (including NHS England aggregate)
  - rtt_icb.csv           ICB-level only (aggregate rows removed)
"""

import re
import sys
from pathlib import Path

import pandas as pd

RAW_DIR = Path(__file__).resolve().parent.parent / "data" / "raw"
OUT_DIR = Path(__file__).resolve().parent.parent / "data" / "processed"

MONTH_MAP = {
    "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
    "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
    "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12",
}

HEADER_ROW = 13  # 0-indexed; row 14 in Excel


def parse_period(filename: str) -> str:
    """Extract YYYY-MM from filename like 'Incomplete-Commissioner-Apr24-...'."""
    m = re.search(r"Commissioner-([A-Za-z]{3})(\d{2})-", filename)
    if not m:
        raise ValueError(f"Cannot parse period from: {filename}")
    month = MONTH_MAP[m.group(1).capitalize()]
    year = f"20{m.group(2)}"
    return f"{year}-{month}"


def read_sheet(filepath: Path, sheet_name: str) -> pd.DataFrame:
    """Read one sheet, skip the 13-row metadata header, drop blank col A."""
    df = pd.read_excel(
        filepath,
        sheet_name=sheet_name,
        header=HEADER_ROW,
        engine="openpyxl",
    )

    if df.columns[0] is None or str(df.columns[0]).startswith("Unnamed"):
        df = df.iloc[:, 1:]

    df = df.dropna(subset=[df.columns[0]])
    df.columns = [str(c).strip() for c in df.columns]

    period = parse_period(filepath.name)
    df.insert(0, "Period", period)

    return df


def add_pct_52_plus(df: pd.DataFrame) -> pd.DataFrame:
    """Derive % 52 plus weeks = Total 52 plus / Total incomplete pathways."""
    total_col = "Total number of incomplete pathways"
    w52_col = "Total 52 plus weeks"
    if total_col in df.columns and w52_col in df.columns:
        df["% 52 plus weeks"] = df[w52_col] / df[total_col]
    return df


def main() -> None:
    files = sorted(
        [f for f in RAW_DIR.iterdir() if f.name.startswith("Incomplete-Commissioner")],
        key=lambda f: parse_period(f.name),
    )

    if not files:
        print("No raw files found in", RAW_DIR, file=sys.stderr)
        sys.exit(1)

    print(f"Found {len(files)} files to process\n")

    national_frames = []
    icb_frames = []

    for i, f in enumerate(files, 1):
        period = parse_period(f.name)
        print(f"  [{i:2d}/{len(files)}] {period}  {f.name}")

        national_frames.append(read_sheet(f, "National"))
        icb_frames.append(read_sheet(f, "ICB"))

    national = add_pct_52_plus(pd.concat(national_frames, ignore_index=True))
    icb_all = add_pct_52_plus(pd.concat(icb_frames, ignore_index=True))

    # The ICB sheet carries a per-file aggregate row (ICB Code "-", named
    # "NHS ENGLAND"). That row is NHS England's *direct* (specialised-
    # commissioning) activity — roughly 127k pathways — NOT the national total.
    # Drop it for the ICB-level file...
    icb = icb_all[icb_all["ICB Code"] != "-"].copy()

    for df in [national, icb_all, icb]:
        df["Treatment Function"] = df["Treatment Function"].str.replace(
            r"^Other - ", "", regex=True
        )

    # ...and rebuild the stacked file as ICB rows + a TRUE national row set taken
    # from the National sheet (the all-England totals, ~7.0M pathways, with the
    # correct median / 92nd-percentile waits that cannot be recovered by summing
    # ICBs). Labelled "NHS ENGLAND" / code "-" so the app's national() helper
    # keeps selecting it unchanged.
    national_rows = national.copy()
    national_rows.insert(1, "ICB Code", "-")
    national_rows.insert(2, "ICB Name", "NHS ENGLAND")
    national_rows = national_rows[icb.columns]  # align column order
    icb_stacked = pd.concat([icb, national_rows], ignore_index=True)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    national.to_csv(OUT_DIR / "rtt_national.csv", index=False)
    icb_stacked.to_csv(OUT_DIR / "rtt_icb_stacked.csv", index=False)
    icb.to_csv(OUT_DIR / "rtt_icb.csv", index=False)

    periods = national["Period"].nunique()
    p_min, p_max = national["Period"].min(), national["Period"].max()
    icb_count = icb["ICB Code"].nunique()

    print(f"\nOutput written to {OUT_DIR}/")
    print(f"  rtt_national.csv      {len(national):>8,} rows")
    print(f"  rtt_icb_stacked.csv   {len(icb_stacked):>8,} rows")
    print(f"  rtt_icb.csv           {len(icb):>8,} rows  ({icb_count} ICBs)")
    print(f"  Periods: {periods} months ({p_min} to {p_max})")


if __name__ == "__main__":
    main()
