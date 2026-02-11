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
from typing import Any

import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import OpenCage
from geopy.extra.rate_limiter import RateLimiter
from geopy.exc import GeocoderServiceError, GeocoderTimedOut
from dotenv import load_dotenv
import ssl

try:
    import certifi
except ImportError:
    certifi = None

# Error tokens we may write into the distance column
ERR_MISSING_COLUMNS = "ERROR_MISSING_COLUMNS"
ERR_MISSING_VALUE = "ERROR_MISSING_VALUE"
ERR_GEOCODE_FAIL = "ERROR_GEOCODE_FAIL"
ERR_LOOKUP_EXCEPTION = "ERROR_LOOKUP_EXCEPTION"

# Known misspellings/aliases to improve geocoding hit rate.
COMMON_CORRECTIONS = {
    "chandigardh": "chandigarh",
    "mumabi": "mumbai",
}

# Tokens that should be treated as missing and not sent to the geocoder.
INVALID_TOKENS = {"", ".", "-", "n/a", "na", "nan", "none", "null", " "}


@dataclass
class GeocodeBias:
    label: str
    kwargs: dict[str, Any]
    low_threshold_km: float | None = None
    high_threshold_km: float | None = None


@dataclass
class Config:
    input_path: str
    output_path: str
    start_col: str = "Starting"
    dest_col: str = "Destination"
    distance_col: str = "Distance_km"
    user_agent: str = "perigon-distance-calculator"
    api_key: str | None = None
    cache_file: str = ".distance_cache.json"
    only_sheets: Tuple[str, ...] = ()
    debug_geocode: bool = False
    min_delay: float = 1.0
    max_retries: int = 2
    error_wait: float = 2.0


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
        help="Custom user agent string for OpenCage",
    )
    parser.add_argument(
        "--api-key",
        help="API key for OpenCage (read from OPENCAGE_API_KEY env var if not provided)",
    )
    parser.add_argument(
        "--cache-file",
        default=".distance_cache.json",
        help="Path to persistent geocode cache file (default: .distance_cache.json in CWD)",
    )
    parser.add_argument(
        "--only-sheets",
        help="Comma-separated list of sheet names to process (default: all)",
    )
    parser.add_argument(
        "--debug-geocode",
        action="store_true",
        help="Print each geocode request/response for tracing (verbose)",
    )
    parser.add_argument(
        "--min-delay",
        type=float,
        default=1.0,
        help="Minimum seconds between geocode requests (default: 1.0, be polite to OSM)",
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=2,
        help="Retries for geocode call on errors/timeouts (default: 2)",
    )
    parser.add_argument(
        "--error-wait",
        type=float,
        default=2.0,
        help="Seconds to wait between retries after an error (default: 2.0)",
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
        api_key=args.api_key,
        cache_file=os.path.abspath(args.cache_file),
        only_sheets=tuple(s.strip() for s in args.only_sheets.split(",")) if args.only_sheets else (),
        debug_geocode=args.debug_geocode,
        min_delay=args.min_delay,
        max_retries=args.max_retries,
        error_wait=args.error_wait,
    )


def geocode_city(
    name: Any,
    geocode_fn,
    cache: dict[str, tuple[float, float]],
    debug: bool,
    geocode_kwargs: dict[str, Any] | None = None,
):
    """Return (lat, lon) for a city name or raise/return None."""
    if not isinstance(name, str) or not name.strip():
        return None

    raw = name.strip()
    key_normalized = raw.lower()
    if key_normalized in INVALID_TOKENS:
        if debug:
            print(f"         [skip] '{raw}' treated as missing")
        return None
    corrected = COMMON_CORRECTIONS.get(key_normalized, raw)

    # Build a cache key that includes provider/bias context to avoid reusing
    # results from runs with different country bias.
    cache_key_parts = [corrected.lower()]
    if geocode_kwargs:
        for k, v in sorted(geocode_kwargs.items()):
            cache_key_parts.append(f"{k}={v}")
    cache_key = "|".join(cache_key_parts)

    if corrected != raw and debug:
        print(f"         [normalize] '{raw}' -> '{corrected}'")

    if cache_key in cache:
        if debug:
            print(f"         [cache hit] '{corrected}' -> {cache[cache_key]}")
        return cache[cache_key]

    if debug:
        print(f"         [geocode] '{corrected}' ...")
    try:
        location = geocode_fn(corrected, **(geocode_kwargs or {}))
    except Exception as exc:
        if debug:
            print(f"         [geocode] '{corrected}' exception: {exc}")
        return None
    if location is None:
        if debug:
            print(f"         [geocode] '{corrected}' -> None")
        return None

    coords = (location.latitude, location.longitude)
    cache[cache_key] = coords
    if debug:
        print(f"         [geocode] '{corrected}' -> {coords}")
    return coords


def compute_distance_pair(
    row,
    cfg: Config,
    geocode_fn,
    cache: dict[str, tuple[float, float]],
    bias_plan: tuple[GeocodeBias, ...],
):
    """
    Try geocoding both endpoints under each bias in order.
    bias_plan: sequence of GeocodeBias entries
    Returns (result, flag):
      result: distance float or error token
      flag: "" or diagnostic string
    """
    start = row.get(cfg.start_col)
    dest = row.get(cfg.dest_col)

    if not isinstance(start, str) or not isinstance(dest, str) or not start.strip() or not dest.strip():
        return ERR_MISSING_VALUE, "MISSING_VALUE"

    for bias in bias_plan:
        coords1 = geocode_city(start, geocode_fn, cache, cfg.debug_geocode, bias.kwargs)
        coords2 = geocode_city(dest, geocode_fn, cache, cfg.debug_geocode, bias.kwargs)
        if coords1 and coords2:
            dist = round(geodesic(coords1, coords2).kilometers, 3)
            if bias.low_threshold_km and dist < bias.low_threshold_km:
                if cfg.debug_geocode:
                    print(f"         [flag {bias.label}] {start}->{dest} = {dist} km below {bias.low_threshold_km}")
                return dist, "CHECK_DISTANCE_LOW"
            if bias.high_threshold_km and dist > bias.high_threshold_km:
                if cfg.debug_geocode:
                    print(f"         [flag {bias.label}] {start}->{dest} = {dist} km exceeds {bias.high_threshold_km}")
                return dist, "CHECK_DISTANCE_HIGH"
            if cfg.debug_geocode:
                print(f"         [distance {bias.label}] {start} -> {dest}: {dist} km")
            return dist, ""
    return ERR_GEOCODE_FAIL, "GEOCODE_FAIL"


def build_geocoder(cfg: Config):
    """Return a rate-limited geocode callable using OpenCage only."""
    ssl_context = build_ssl_context()
    geocoder = OpenCage(
        api_key=cfg.api_key,
        timeout=5,
        user_agent=cfg.user_agent,
        ssl_context=ssl_context,
    )
    return RateLimiter(
        geocoder.geocode,
        min_delay_seconds=cfg.min_delay,
        max_retries=cfg.max_retries,
        error_wait_seconds=cfg.error_wait,
        swallow_exceptions=True,  # return None instead of raising on HTTP errors
    )


def build_ssl_context():
    if certifi:
        return ssl.create_default_context(cafile=certifi.where())
    return ssl.create_default_context()


def process_workbook(cfg: Config) -> None:
    # Ensure output path directory is writable before heavy work.
    out_dir = Path(cfg.output_path).parent
    if not os.access(out_dir, os.W_OK):
        sys.exit(f"Output directory is not writable: {out_dir}")

    cache: dict[str, tuple[float, float]] = load_cache(cfg.cache_file)

    print(f"[1/4] Reading workbook: {cfg.input_path}")
    sheets = pd.read_excel(cfg.input_path, sheet_name=None)
    print(f"[2/4] Loaded {len(sheets)} sheet(s): {', '.join(sheets.keys())}")

    if cfg.only_sheets:
        missing = [s for s in cfg.only_sheets if s not in sheets]
        if missing:
            print(f"Warning: requested --only-sheets not found: {', '.join(missing)}")

    geocode_fn = build_geocoder(cfg)

    output_sheets = {}
    for sheet_name, df in sheets.items():
        if cfg.only_sheets and sheet_name not in cfg.only_sheets:
            print(f"       Skipped '{sheet_name}' (not in --only-sheets)")
            output_sheets[sheet_name] = df
            continue

        df_out = df.copy()
        if cfg.start_col in df_out.columns and cfg.dest_col in df_out.columns:
            total_rows = len(df_out)
            # For debug runs, cap processing to avoid long waits
            debug_row_limit = 50 if cfg.debug_geocode else None
            df_proc = df_out if not debug_row_limit else df_out.head(debug_row_limit)
            proc_rows = len(df_proc)
            print(f"[3/4] Processing sheet '{sheet_name}' with {proc_rows}/{total_rows} row(s){' (debug limited)' if debug_row_limit else ''}")

            if cfg.distance_col in df_out.columns:
                print(f"Warning: sheet '{sheet_name}' already has column '{cfg.distance_col}', values will be overwritten.")

            # Sheet-specific geocode bias and outlier thresholds
            bias_plan: tuple[GeocodeBias, ...]
            if sheet_name.lower() == "uniquetrains":
                bias_plan = (
                    GeocodeBias("GB", {"country": "gb"}, None, 800.0),
                    GeocodeBias("IN", {"country": "in"}, None, 2000.0),
                    GeocodeBias("EU", {}, None, 2000.0),
                )
            elif sheet_name.lower() == "uniqueflights":
                bias_plan = (
                    GeocodeBias("ANY", {}, 200.0, 8000.0),
                    GeocodeBias("IN", {"country": "in"}, 200.0, 8000.0),
                    GeocodeBias("GB", {"country": "gb"}, 200.0, 8000.0),
                )
            else:
                bias_plan = (GeocodeBias("ANY", {}, None, None),)

            distances = ["DEBUG_SKIPPED" for _ in range(total_rows)] if debug_row_limit else []
            flags = ["DEBUG_SKIPPED" for _ in range(total_rows)] if debug_row_limit else []

            for idx, (_, row) in enumerate(df_proc.iterrows(), start=1):
                result, flag = compute_distance_pair(row, cfg, geocode_fn, cache, bias_plan)
                if debug_row_limit:
                    distances[idx - 1] = result
                    flags[idx - 1] = flag
                else:
                    distances.append(result)
                    flags.append(flag)
                if idx % 25 == 0 or idx == proc_rows:
                    print(f"       {sheet_name}: {idx}/{proc_rows} rows complete")
                    # brief pause so print buffer flushes in some terminals
                    time.sleep(0.01)
                if cfg.debug_geocode:
                    print(f"       [row {idx}] start='{row.get(cfg.start_col)}' dest='{row.get(cfg.dest_col)}' -> {result}")
            df_out[cfg.distance_col] = distances
            if sheet_name.lower() in ("uniquetrains", "uniqueflights"):
                df_out["Distance_flag"] = flags
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
    # Load environment variables from .env if present
    load_dotenv()
    cfg = parse_args()
    # pick up API key from env if not provided
    if not cfg.api_key:
        cfg.api_key = os.getenv("OPENCAGE_API_KEY")
    if not cfg.api_key:
        sys.exit("OPENCAGE_API_KEY not set. Provide --api-key or set it in the environment/.env.")
    try:
        process_workbook(cfg)
    except KeyboardInterrupt:
        print("\nProcess interrupted by user. No output file written.")


def load_cache(cache_path: str) -> dict[str, tuple[float, float]]:
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


def save_cache(cache_path: str, cache: dict[str, tuple[float, float]]) -> None:
    try:
        path = Path(cache_path)
        with path.open("w", encoding="utf-8") as f:
            json.dump(cache, f)
    except Exception:
        # Cache write failures should not break main flow
        pass


if __name__ == "__main__":
    main()
