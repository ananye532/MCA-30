# Daily log automation

`mca31.xlsx` (Daily Log sheet) gets one new row per day, committed to this repo.

## Scripts

| Script | What it does |
| --- | --- |
| `tools/generate_entry.py` | Samples one day's values from the ranges already in the workbook and prints/writes them as JSON. The numbers are synthesized, not measured: durations are drawn around the recent median within the observed min/max, class minutes and class count come from the same weekday's last four entries, and mood/energy are drawn from their historical frequencies. The seed is the date, so the same date always yields the same entry. |
| `tools/append_daily_row.py` | Writes one JSON entry (or `--flag` values) into the next free row of the Daily Log, keeping the sheet's formulas in columns I/J, its cell styles and its dropdown validations. Re-running for a date already in the sheet overwrites that row instead of appending a duplicate. |
| `tools/daily_update.sh` | Generate → write → commit → push, with push retries. Idempotent: running it twice for the same date commits nothing the second time. |

## Daily run

```bash
tools/daily_update.sh              # today
tools/daily_update.sh 2026-09-21   # a specific date
```

To use real values instead of generated ones, drop them in first and the
generator is skipped:

```bash
cat > pending/2026-09-21.json <<'JSON'
{"date": "2026-09-21", "sleep": 465, "fitness": 40, "study": 120, "coding": 55,
 "class_min": 250, "classes": 5, "other": 60, "feeling": "Good",
 "satisfaction": "Satisfied", "energy": "High", "notes": "..."}
JSON
tools/daily_update.sh 2026-09-21
```

`pending/<date>.json` is committed alongside the workbook so each commit has a
readable diff — the `.xlsx` itself is binary.

## Schedule

A Claude Code Routine fires daily at 21:17 IST (15:47 UTC) in a fresh session,
which clones this repo and runs `tools/daily_update.sh`. Requirements: Python 3
with `openpyxl` (the script installs it if missing) and push access to
`ananye532/MCA-30`.

Column rules (minutes for every duration, allowed dropdown values, which cells
are formula-driven) are in the workbook's own `Instructions` sheet.
