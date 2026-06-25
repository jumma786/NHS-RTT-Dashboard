"""Geography helper: load ICB boundaries and join them to the RTT data by name.

The ONS BSC boundary file carries ICB names (ICB23NM) and ONS codes (ICB23CD)
but not the 3-char NHS org code used in the RTT data, so we match on a
normalised version of the ICB name (all 42 align exactly).
"""
from __future__ import annotations

import json
import os
from functools import lru_cache

import data as D

_GEOJSON = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "geo", "icb.geojson")
GEO_KEY = "ICB23NM"  # property used as the choropleth feature id


def _norm(s: str) -> str:
    return (
        s.upper()
        .replace("NHS ", "")
        .replace(" INTEGRATED CARE BOARD", "")
        .replace(" ICB", "")
        .replace(",", "")
        .replace("AND", "&")
        .strip()
    )


@lru_cache(maxsize=None)
def geojson() -> dict:
    with open(_GEOJSON, encoding="utf-8") as fh:
        return json.load(fh)


@lru_cache(maxsize=None)
def _name_map() -> dict:
    """Map data ICB Name -> boundary ICB23NM via normalised match."""
    geo_by_norm = {_norm(f["properties"][GEO_KEY]): f["properties"][GEO_KEY] for f in geojson()["features"]}
    out = {}
    for nm in D.icbs()[D.ICB_NAME]:
        out[nm] = geo_by_norm.get(_norm(nm))
    return out


def attach(df):
    """Add a ``geo_name`` column (matching boundary ICB23NM) to a snapshot frame."""
    df = df.copy()
    df["geo_name"] = df[D.ICB_NAME].map(_name_map())
    return df
