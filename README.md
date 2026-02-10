# Distance Calculator CLI

Add great-circle travel distances to every worksheet in an Excel workbook using `geopy`. Supports public Nominatim (not recommended for batches) and OpenCage with an API key.

## Prerequisites
- Python 3.9+
- Excel file with columns `Starting` and `Destination` (default names; customizable via flags).

## Setup
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
```

Create a `.env` file (already gitignored) and add your OpenCage key if you’ll use that provider:
```
OPENCAGE_API_KEY=your_real_key_here
```

## Usage
```bash
python main.py "/path/to/Annual Travel Data 2025.xlsx"
```

Behavior:
- Reads all sheets in the workbook.
- If a sheet has `Starting` and `Destination`, adds `Distance_km` with geodesic distance; otherwise fills that column with `ERROR_MISSING_COLUMNS`.
- Geocode failures write `ERROR_GEOCODE_FAIL`; empty cells write `ERROR_MISSING_VALUE`.
- Output is saved beside the input as `New_<original filename>.xlsx`.

### Optional flags
- `--provider` choose `opencage` (recommended for batches) or `nominatim` (default; public endpoint can 403).
- `--api-key` API key for provider (falls back to `OPENCAGE_API_KEY` env var).
- `--start-col` Override origin column name (default: `Starting`).
- `--dest-col`  Override destination column name (default: `Destination`).
- `--distance-col` Output column name (default: `Distance_km`).
- `--user-agent` Custom Nominatim user agent (default: `perigon-distance-calculator`).
- `--only-sheets` Comma-separated list of sheet names to process.
- `--debug-geocode` Verbose geocode logging.
- `--min-delay`, `--max-retries`, `--error-wait` Tune rate limiting/backoff.

### Example
```bash
# Using OpenCage with key from .env
python main.py "./Annual Travel Data 2025.xlsx" \
  --provider opencage \
  --only-sheets UniqueFlights \
  --debug-geocode
```

## Notes
- Public Nominatim is not intended for bulk jobs and may return 403 immediately; use `opencage` with an API key for reliability.
- The original file is never modified; the new file is written alongside it.
