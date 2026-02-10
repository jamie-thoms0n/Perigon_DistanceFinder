# Distance Calculator CLI

Add great-circle travel distances to every worksheet in an Excel workbook using OpenStreetMap (Nominatim) via `geopy`.

## Prerequisites
- Python 3.9+
- Excel file with columns `Starting` and `Destination` (default names; customizable via flags).

## Setup
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
```

## Usage
```bash
python main.py "/Users/jamiethomson/Downloads/Annual Travel Data - 2025.xlsx"
```

Behavior:
- Reads all sheets in the workbook.
- If a sheet has `Starting` and `Destination`, adds `Distance_km` with geodesic distance; otherwise fills that column with `ERROR_MISSING_COLUMNS`.
- Geocode failures write `ERROR_GEOCODE_FAIL`; empty cells write `ERROR_MISSING_VALUE`.
- Output is saved beside the input as `New_<original filename>.xlsx`.

### Optional flags
- `--start-col` Override origin column name (default: `Starting`).
- `--dest-col`  Override destination column name (default: `Destination`).
- `--distance-col` Output column name (default: `Distance_km`).
- `--user-agent` Custom Nominatim user agent (default: `perigon-distance-calculator`).

### Example
```bash
python main.py "./Annual Travel Data - 2025.xlsx"
```

## Notes
- Be courteous to OpenStreetMap Nominatim usage policy. Heavy runs should add throttling; the script already caches lookups in-memory per run.
- The original file is never modified; the new file is written alongside it.
