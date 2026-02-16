# Distance Calculator Excel Tool

This project reads every worksheet in an Excel workbook, geocodes the origin and destination cities, then adds a `Distance_km` column with the great-circle distance. It uses the OpenCage Geocoding API exclusively--an API key is required (free tier is fine).

> WARNING: IMPORTANT: Your workbook must contain a column for the origin city and a column for the destination city. By default they must be named exactly `Starting` and `Destination` (capitalized, no spaces). You can override the names with flags if your sheet uses different headers.

The guide below is written for newcomers--no command-line experience assumed.

---

## What You Need

1) **A computer** running macOS or Windows.  
2) **Reliable internet** (the script geocodes each row).  
3) **Python 3.9+** installed (we'll show how).  
4) **OpenCage API key (required)** saved as `OPENCAGE_API_KEY`. Sign up free at https://opencagedata.com/.  
5) **Your Excel workbook** (`.xlsx`), first row as headers and at least the origin/destination columns.

---

## Step 1: Get Python

### macOS
- Open **Spotlight** (Cmd + Space), type `Terminal`, press Enter.
- Check Python: `python3 --version` then Enter.  
  - If you see a version (e.g., `Python 3.12.2`), you're set.  
  - If "command not found", install Homebrew (https://brew.sh), then run:  
    - `brew install python`

### Windows
- Press **Start**, type `cmd`, press Enter to open Command Prompt.
- Check Python: `py --version` then Enter.  
  - If you see a version, you're set.  
  - If not, install Python from https://www.python.org/downloads/ (tick "Add Python to PATH" during install). Reopen Command Prompt and try `py --version` again.

---

## Step 2: Download / Open the Project Folder

Put the project folder (contains `main.py`, `requirements.txt`, `.env`) somewhere easy, e.g., Desktop:
- macOS example path: `/Users/yourname/Desktop/Perigon_DistanceFinder`
- Windows example path: `C:\Users\yourname\Desktop\Perigon_DistanceFinder`

### How to find the path of your folder
- macOS Finder: open Terminal, type `cd ` (trailing space), then drag the folder into the Terminal window and press Enter.  
- Windows File Explorer: click the address bar while inside the folder, copy the full path.

---

## Step 3: Open a Terminal/Command Prompt in the Project Folder

### macOS
```
cd /Users/yourname/Desktop/Perigon_DistanceFinder
```

### Windows
```
cd C:\Users\yourname\Desktop\Perigon_DistanceFinder
```

Tip: `ls` (macOS) or `dir` (Windows) should show `main.py`, `.env`, `requirements.txt`, etc.

---

## Step 4: Install the Required Python Packages

### macOS
```
python3 -m pip install -r requirements.txt
```
If you see a "permission denied" message, try:
```
python3 -m pip install --user -r requirements.txt
```

### Windows
```
py -m pip install -r requirements.txt
```
If you see a permissions message, right-click Command Prompt, choose "Run as administrator", then rerun the command.

> Optional but nice: create a virtual environment first (`python3 -m venv .venv` on macOS or `py -m venv .venv` on Windows, then activate it) to keep dependencies isolated.

---

## Step 5: Add Your API Key to `.env`

Open the `.env` file in a text editor (TextEdit on macOS, Notepad on Windows). Fill in:
```
OPENCAGE_API_KEY=your_real_key_here
```
Save the file. It's ignored by git. If you share the project and don't want to share your key, blank it first.

No key, no run: the script exits if `OPENCAGE_API_KEY` isn't set or passed via `--api-key`.

---

## Step 6: Prepare Your Excel File

- Ensure the sheet has columns named `Starting` and `Destination` (exact spelling and capitalization). You can use different names with flags `--start-col` and `--dest-col` when running.  
- Keep city names clean: avoid empty cells; each row needs both values.  
- Save as `.xlsx`. Note the full path to the file, e.g.:  
  - macOS: `/Users/yourname/Desktop/Annual Travel Data 2025.xlsx`  
  - Windows: `C:\Users\yourname\Desktop\Annual Travel Data 2025.xlsx`

---

## Step 7: Run the Script

### macOS
```
python3 main.py "/Users/yourname/Desktop/Annual Travel Data 2025.xlsx"
```

### Windows
```
py main.py "C:\\Users\\yourname\\Desktop\\Annual Travel Data 2025.xlsx"
```

What happens:
- Reads every sheet in the workbook.  
- If a sheet has the origin/destination columns, it adds `Distance_km` with great-circle distance (km).  
- Sheets missing the columns get `ERROR_MISSING_COLUMNS` in that field so you can spot them.  
- Writes a new workbook next to the input named `New_<original filename>.xlsx` (the original stays untouched).  
- Creates a geocode cache `.distance_cache.json` to speed up reruns.

Example with custom options (macOS):
```
python3 main.py "/Users/yourname/Desktop/Travel.xlsx" \
  --start-col Origin \
  --dest-col DestinationCity \
  --only-sheets UniqueFlights,UniqueTrains \
  --debug-geocode
```

---

## Understanding the Output

Columns added/used:
- `Distance_km` -- great-circle distance between the two cities (kilometers).
- `Distance_flag` -- only on sheets named `UniqueTrains` or `UniqueFlights`; marks rows that look unusually near/far and may need checking.
- Error tokens: `ERROR_MISSING_COLUMNS`, `ERROR_MISSING_VALUE`, `ERROR_GEOCODE_FAIL`, `ERROR_LOOKUP_EXCEPTION` so you know why a row couldn't be processed.

The new file appears beside your input, e.g., input `/Users/you/Desktop/Trips.xlsx` produces `/Users/you/Desktop/New_Trips.xlsx`.

---

## How to Know It's Working

- Terminal shows steps like `[1/4] Reading workbook`, row counts per sheet, then `[4/4] Writing output workbook ...`  
- For every 25 rows you'll see progress updates; `--debug-geocode` prints each lookup.  
- At the end: `Processed X sheet(s). Output written to: ...`

---

## Common Issues & Fixes

- **"python: command not found" (macOS)** -- Install Python via Homebrew (`brew install python`), then rerun using `python3`.
- **"py is not recognized" (Windows)** -- Reinstall Python and tick "Add Python to PATH", then use `py`.
- **403 / lots of failures** -- Usually an API quota issue. Verify `OPENCAGE_API_KEY` and wait/retry if you've hit the free-tier limit.
- **"Input file not found"** -- Double-check the full path you passed into the command (quotes help when paths have spaces).
- **Columns not detected** -- Ensure headers match exactly (`Starting`, `Destination`) or pass `--start-col` / `--dest-col`.
- **Very long runs** -- Geocoding is rate-limited to be polite to the API; keep the terminal open and let it finish.

---

## Optional: Custom Settings

- `--api-key` to pass the OpenCage key via CLI instead of `.env`.  
- `--start-col`, `--dest-col`, `--distance-col` to rename columns.  
- `--only-sheets Sheet1,Sheet2` to process specific tabs.  
- `--cache-file` to choose a custom cache path.  
- `--min-delay`, `--max-retries`, `--error-wait` to tune rate limits/backoff.  
- `--user-agent` to set a custom user agent for geocoding.  
- `--debug-geocode` to print every geocode attempt.

---

## Quick Reference (Cheat Sheet)

1) `cd /path/to/Perigon_DistanceFinder`  
2) Install deps: `python3 -m pip install -r requirements.txt` (mac) or `py -m pip install -r requirements.txt` (win).  
3) Fill `.env` with `OPENCAGE_API_KEY`.  
4) Run:  
   - macOS: `python3 main.py "/path/to/YourFile.xlsx"`  
   - Windows: `py main.py "C:\\path\\to\\YourFile.xlsx"`  
5) Find results next to the input as `New_<filename>.xlsx`.

---

## Where Files Go

- **Input**: anywhere you like; you provide the full path.  
- **Output**: saved beside your input as `New_<original name>.xlsx`.  
- **Cache**: `.distance_cache.json` in the project folder (safe to delete if it grows; it will be recreated).  
- **Config/keys**: `.env` in the project folder (ignored by git).

---

## Credits & Copyright
Developed by Jamie Thomson (Jamie.wlt@outlook.com). Please reach out with any problems.
Copyright © Perigon Partners. All rights reserved.
