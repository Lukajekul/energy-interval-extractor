# Energy Interval Extractor

Parses raw 15-minute interval electricity meter readings (semicolon-delimited CSV exports) and pulls out the fields that actually matter — timestamp, energy value, direction, and method — into a clean Excel workbook, one sheet per measurement point.

## What it does

- Reads one or more raw `.csv` meter export files (semicolon-delimited, first line is a header).
- For each 15-minute reading:
  - Converts the timestamp from UTC to local time.
  - Computes the interval's "To" time (from + 15 minutes).
  - Converts the kWh value from its raw number format (dots used as thousand separators, no decimal point) into a normal float.
  - Records the reading's direction (`r` → Received, `s` → Sent).
  - Counts how many readings use a method other than `L3`, and flags it as a warning.
- Sorts all readings by their "From" time.
- Writes everything into a new sheet in a single `.xlsx` workbook, named `MP_<last 3 digits of the measurement point ID>`.
- Repeats for each input file, all saved into the same output workbook.

## Input format

Each row of the source CSV is expected to look like:

```
ID;MeasurementPointID;From;Value;Direction;Method;Created;Modified
```

| Column | Meaning |
|---|---|
| ID | Unique row identifier (not used in output) |
| MeasurementPointID | Meter/measurement point ID — last 3 digits become the sheet name |
| From | Start of the 15-minute interval, UTC, format `YYYY-MM-DD HH:MM:SS.fff` |
| Value | Raw kWh reading, e.g. `148.000.000` (dots are thousand separators; divide by 1,000,000 for the actual kWh value) |
| Direction | `r` (received) or `s` (sent) |
| Method | Phase/method code, e.g. `L1`, `L2`, `L3`, `01`–`04` |
| Created / Modified | Extra metadata fields (not used by the script) |

The first line of the file is treated as a header and skipped.

## Output format

One `.xlsx` file, with one sheet per measurement point:

| From | To | MeasurmentHethod | Value in KWH |
|---|---|---|---|
| 2026-06-01 00:00 | 2026-06-01 00:15 | 01 | 148 |
| 2026-06-01 00:15 | 2026-06-01 00:30 | L1 | 523 |
| ... | ... | ... | ... |

Off to the side (columns F/G) it also writes:
- The measurement point ID
- The direction (`Recieved`/`Sent`)
- A warning message listing how many readings used a method other than `L3` (only added if at least one exists)

## Requirements

- Python 3.8+
- [`openpyxl`](https://pypi.org/project/openpyxl/)
- [`python-dateutil`](https://pypi.org/project/python-dateutil/)

## Installation

```bash
git clone <your-repo-url>
cd energy-interval-extractor
pip install openpyxl python-dateutil
```

## Usage

```bash
python energy_interval_extractor.py
```

You'll be prompted for:
1. How many files you're processing
2. The path to each file
3. What to name the resulting `.xlsx` file

## Examples

See the `examples/` folder:

- `example_input.csv` — a sample raw input file (10 rows, sequential 15-minute intervals)
- `example_output.xlsx` — the actual output produced by running the script against `example_input.csv`
- `generate_test_data.py` — generates random sample input files in the same format, useful for testing changes to the script

## Key variables (for anyone reading the code)

- `rawRows` — all lines from the input file, split by `;`
- `sheet` — the worksheet currently being written to
- `fullList` — all processed rows for the current file, before sorting
- `warningDictionary` — counts of readings by method, for anything that isn't `L3`
- `measurementPointId` — the measurement point ID, captured from the first data row
- `kwhValue` — the converted (real) kWh value
- `directionValue` — `"Recieved"` or `"Sent"`, written once per sheet

## License

See [LICENSE](LICENSE).
