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
import ssl
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Tuple, Any

import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim, OpenCage
from geopy.extra.rate_limiter import RateLimiter
from geopy.exc import GeocoderServiceError, GeocoderTimedOut
from dotenv import load_dotenv

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


@dataclass
class Config:
    input_path: str
    output_path: str
    start_col: str = "Starting"
    dest_col: str = "Destination"
    distance_col: str = "Distance_km"
    user_agent: str = "jxxthomson@gmail.com"
    provider: str = "nominatim"  # or "opencage"
    api_key: str | None = None
    cache_file: str = ".distance_cache.json"
    only_sheets: Tuple[str, ...] = ()
    debug_geocode: bool = False
    min_delay: float = 1.0
    max_retries: int = 2
    error_wait: float = 2.0
    insecure_ssl: bool = False


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
        "--provider",
        choices=["nominatim", "opencage"],
        default="nominatim",
        help="Geocoding provider (default: nominatim; use opencage with API key for reliability)",
    )
    parser.add_argument(
        "--api-key",
        help="API key for provider (read from OPENCAGE_API_KEY env var if not provided)",
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
    parser.add_argument(
        "--insecure-ssl",
        action="store_true",
        help="Skip SSL certificate verification for geocoding (not recommended).",
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
        provider=args.provider,
        api_key=args.api_key,
        cache_file=os.path.abspath(args.cache_file),
        only_sheets=tuple(s.strip() for s in args.only_sheets.split(",")) if args.only_sheets else (),
        debug_geocode=args.debug_geocode,
        min_delay=args.min_delay,
        max_retries=args.max_retries,
        error_wait=args.error_wait,
        insecure_ssl=args.insecure_ssl,
    )


def geocode_city(name: Any, geocode_fn, cache: Dict[str, Tuple[float, float]], debug: bool):
    """Return (lat, lon) for a city name or raise/return None."""
    if not isinstance(name, str) or not name.strip():
        return None

    raw = name.strip()
    key = raw.lower()
    corrected = COMMON_CORRECTIONS.get(key, raw)
    key_cache = corrected.lower()

    if corrected != raw and debug:
        print(f"         [normalize] '{raw}' -> '{corrected}'")

    if key_cache in cache:
        if debug:
            print(f"         [cache hit] '{corrected}' -> {cache[key_cache]}")
        return cache[key_cache]

    if debug:
        print(f"         [geocode] '{corrected}' ...")
    try:
        location = geocode_fn(corrected)
    except Exception as exc:
        if debug:
            print(f"         [geocode] '{corrected}' exception: {exc}")
        return None
    if location is None:
        if debug:
            print(f"         [geocode] '{corrected}' -> None")
        return None

    coords = (location.latitude, location.longitude)
    cache[key_cache] = coords
    if debug:
        print(f"         [geocode] '{corrected}' -> {coords}")
    return coords


def compute_distance(row, cfg: Config, geocode_fn, cache: Dict[str, Tuple[float, float]]):
    start = row.get(cfg.start_col)
    dest = row.get(cfg.dest_col)

    if not isinstance(start, str) or not isinstance(dest, str) or not start.strip() or not dest.strip():
        return ERR_MISSING_VALUE

    try:
        coords1 = geocode_city(start, geocode_fn, cache, cfg.debug_geocode)
        coords2 = geocode_city(dest, geocode_fn, cache, cfg.debug_geocode)
        if coords1 is None or coords2 is None:
            return ERR_GEOCODE_FAIL
        dist = round(geodesic(coords1, coords2).kilometers, 3)
        if cfg.debug_geocode:
            print(f"         [distance] {start} -> {dest}: {dist} km")
        return dist
    except (GeocoderTimedOut, GeocoderServiceError):
        return ERR_GEOCODE_FAIL
    except Exception:
        return ERR_LOOKUP_EXCEPTION


def build_geocoder(cfg: Config):
    """Return a rate-limited geocode callable based on provider."""
    ssl_context = build_ssl_context(cfg)

    if cfg.provider == "opencage":
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
            swallow_exceptions=False,
        )

    # Default: public Nominatim (may be blocked; requires user_agent and politeness)
    geocoder = Nominatim(user_agent=cfg.user_agent, timeout=5, ssl_context=ssl_context)
    return RateLimiter(
        geocoder.geocode,
        min_delay_seconds=cfg.min_delay,
        max_retries=cfg.max_retries,
        error_wait_seconds=cfg.error_wait,
        swallow_exceptions=False,
    )


def build_ssl_context(cfg: Config):
    """Build an SSL context honoring certifi and --insecure-ssl."""
    if cfg.insecure_ssl:
        return ssl._create_unverified_context()
    if certifi:
        return ssl.create_default_context(cafile=certifi.where())
    return ssl.create_default_context()


def process_workbook(cfg: Config) -> None:
    cache: Dict[str, Tuple[float, float]] = load_cache(cfg.cache_file)

    print(f"[1/4] Reading workbook: {cfg.input_path}")
    sheets = pd.read_excel(cfg.input_path, sheet_name=None)
    print(f"[2/4] Loaded {len(sheets)} sheet(s): {', '.join(sheets.keys())}")

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
            print(f"[3/4] Processing sheet '{sheet_name}' with {total_rows} row(s)")
            distances = []
            for idx, (_, row) in enumerate(df_out.iterrows(), start=1):
                result = compute_distance(row, cfg, geocode_fn, cache)
                distances.append(result)
                if idx % 25 == 0 or idx == total_rows:
                    print(f"       {sheet_name}: {idx}/{total_rows} rows complete")
                    # brief pause so print buffer flushes in some terminals
                    time.sleep(0.01)
                if cfg.debug_geocode:
                    print(f"       [row {idx}] start='{row.get(cfg.start_col)}' dest='{row.get(cfg.dest_col)}' -> {result}")
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
    # Load environment variables from .env if present
    load_dotenv()
    cfg = parse_args()
    # pick up API key from env if not provided
    if cfg.provider == "opencage" and not cfg.api_key:
        cfg.api_key = os.getenv("OPENCAGE_API_KEY")
    if cfg.provider == "opencage" and not cfg.api_key:
        sys.exit("OPENCAGE_API_KEY not set. Provide --api-key or set it in the environment/.env.")
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
