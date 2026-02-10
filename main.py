"""
CLI tool: add great-circle travel distances to every worksheet in an Excel file.

Usage:
    python main.py input.xlsx

Behavior:
- Reads all sheets.
- If a sheet has the expected columns (default: Starting, Destination), adds a
  Distance_km column with geodesic distance between the two locations using
  OpenStreetMap Nominatim via geopy.
- If required columns are missing or a row cannot be geocoded, writes a clear
  error token instead of crashing.
- Writes a new workbook beside the input named New_<original filename>.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Tuple, Any

import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderServiceError, GeocoderTimedOut


# Error tokens we may write into the distance column
ERR_MISSING_COLUMNS = "ERROR_MISSING_COLUMNS"
ERR_MISSING_VALUE = "ERROR_MISSING_VALUE"
ERR_GEOCODE_FAIL = "ERROR_GEOCODE_FAIL"
ERR_LOOKUP_EXCEPTION = "ERROR_LOOKUP_EXCEPTION"


@dataclass
class Config:
    input_path: str
    output_path: str
    start_col: str = "Starting"
    dest_col: str = "Destination"
    distance_col: str = "Distance_km"
    user_agent: str = "perigon-distance-calculator"
    cache_file: str = ".distance_cache.json"


def parse_args() -> Config:
    parser = argparse.ArgumentParser(
        description="Add distance column to all Excel sheets with Starting/Destination columns",
    )
    parser.add_argument("input", help="Path to the input Excel file")
    parser.add_argument(
        "--start-col",
        default="Starting",
        help="Column name for origin city (default: Starting)",
    )
    parser.add_argument(
        "--dest-col",
        default="Destination",
        help="Column name for destination city (default: Destination)",
    )
    parser.add_argument(
        "--distance-col",
        default="Distance_km",
        help="Name of the output distance column (default: Distance_km)",
    )
    parser.add_argument(
        "--user-agent",
        default="perigon-distance-calculator",
        help="Custom user agent for Nominatim (required by OSM policy)",
    )
    parser.add_argument(
        "--cache-file",
        default=".distance_cache.json",
        help="Path to persistent geocode cache file (default: .distance_cache.json in CWD)",
    )

    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    if not os.path.isfile(input_path):
        sys.exit(f"Input file not found: {input_path}")

    dirname, basename = os.path.split(input_path)
    output_path = os.path.join(dirname, f"New_{basename}")

    return Config(
        input_path=input_path,
        output_path=output_path,
        start_col=args.start_col,
        dest_col=args.dest_col,
        distance_col=args.distance_col,
        user_agent=args.user_agent,
        cache_file=os.path.abspath(args.cache_file),
    )


def geocode_city(name: Any, geolocator: Nominatim, cache: Dict[str, Tuple[float, float]]):
    """Return (lat, lon) for a city name or raise/return None."""
    if not isinstance(name, str) or not name.strip():
        return None
    key = name.strip().lower()
    if key in cache:
        return cache[key]
    location = geolocator.geocode(name)
    if location is None:
        return None
    coords = (location.latitude, location.longitude)
    cache[key] = coords
    return coords


def compute_distance(row, cfg: Config, geolocator: Nominatim, cache: Dict[str, Tuple[float, float]]):
    start = row.get(cfg.start_col)
    dest = row.get(cfg.dest_col)

    if not isinstance(start, str) or not isinstance(dest, str) or not start.strip() or not dest.strip():
        return ERR_MISSING_VALUE

    try:
        coords1 = geocode_city(start, geolocator, cache)
        coords2 = geocode_city(dest, geolocator, cache)
        if coords1 is None or coords2 is None:
            return ERR_GEOCODE_FAIL
        return round(geodesic(coords1, coords2).kilometers, 3)
    except (GeocoderTimedOut, GeocoderServiceError):
        return ERR_GEOCODE_FAIL
    except Exception:
        return ERR_LOOKUP_EXCEPTION


def process_workbook(cfg: Config) -> None:
    cache: Dict[str, Tuple[float, float]] = load_cache(cfg.cache_file)

    print(f"[1/4] Reading workbook: {cfg.input_path}")
    sheets = pd.read_excel(cfg.input_path, sheet_name=None)
    print(f"[2/4] Loaded {len(sheets)} sheet(s): {', '.join(sheets.keys())}")

    geolocator = Nominatim(user_agent=cfg.user_agent, timeout=5)

    output_sheets = {}
    for sheet_name, df in sheets.items():
        df_out = df.copy()
        if cfg.start_col in df_out.columns and cfg.dest_col in df_out.columns:
            total_rows = len(df_out)
            print(f"[3/4] Processing sheet '{sheet_name}' with {total_rows} row(s)")
            distances = []
            for idx, (_, row) in enumerate(df_out.iterrows(), start=1):
                distances.append(compute_distance(row, cfg, geolocator, cache))
                if idx % 25 == 0 or idx == total_rows:
                    print(f"       {sheet_name}: {idx}/{total_rows} rows complete")
                    # brief pause so print buffer flushes in some terminals
                    time.sleep(0.01)
            df_out[cfg.distance_col] = distances
            print(f"       Done: {sheet_name}")
        else:
            # Mark the issue so the user sees the problem but keep processing other sheets.
            df_out[cfg.distance_col] = ERR_MISSING_COLUMNS
            print(f"       Skipped '{sheet_name}' (missing columns: {cfg.start_col}, {cfg.dest_col})")
        output_sheets[sheet_name] = df_out

    print(f"[4/4] Writing output workbook: {cfg.output_path}")
    with pd.ExcelWriter(cfg.output_path, engine="openpyxl") as writer:
        for sheet_name, df_out in output_sheets.items():
            df_out.to_excel(writer, sheet_name=sheet_name, index=False)

    save_cache(cfg.cache_file, cache)
    print(f"Processed {len(output_sheets)} sheet(s). Output written to: {cfg.output_path}")


def main():
    cfg = parse_args()
    try:
        process_workbook(cfg)
    except KeyboardInterrupt:
        print("\nProcess interrupted by user. No output file written.")


def load_cache(cache_path: str) -> Dict[str, Tuple[float, float]]:
    path = Path(cache_path)
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
            # ensure tuple conversion for geopy usage
            return {k: tuple(v) for k, v in data.items()}
    except Exception:
        # Corrupt or unreadable cache; start fresh
        return {}


def save_cache(cache_path: str, cache: Dict[str, Tuple[float, float]]) -> None:
    try:
        path = Path(cache_path)
        with path.open("w", encoding="utf-8") as f:
            json.dump(cache, f)
    except Exception:
        # Cache write failures should not break main flow
        pass


if __name__ == "__main__":
    main()
