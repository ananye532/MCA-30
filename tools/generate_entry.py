#!/usr/bin/env python3
"""Synthesize one day's entry from the ranges already present in the workbook.

The values are NOT measured — they are sampled from the historical rows of the
Daily Log so that a generated day sits inside the same ranges as the recorded
ones. Class minutes and class count come from the same weekday's recent
history (weekends have been zero); the other durations are sampled around the
recent median; mood/energy are drawn from their historical frequencies.

The random seed is derived from the date, so re-running for the same date
yields the same entry (the daily job stays idempotent).

Usage:
    python3 tools/generate_entry.py                      # prints JSON for today
    python3 tools/generate_entry.py --date 2026-09-21 --out pending/2026-09-21.json
"""

import argparse
import datetime as dt
import json
import random
import statistics
import sys
from pathlib import Path

import openpyxl

SHEET = "Daily Log"
FIRST_DATA_ROW = 6
LAST_DATA_ROW = 401
HISTORY_DAYS = 28

# 1-based column numbers in the Daily Log sheet.
COLUMNS = {
    "date": 1,
    "sleep": 2,
    "fitness": 3,
    "study": 4,
    "coding": 5,
    "class_min": 6,
    "classes": 7,
    "other": 8,
    "feeling": 11,
    "satisfaction": 12,
    "energy": 13,
    "notes": 14,
}

DURATIONS = ["sleep", "fitness", "study", "coding", "other"]

# Hard ceilings for any generated row.
MAX_CLASSES = 8
MAX_SLEEP = 599  # keep sleep under 600 minutes
MAX_FITNESS = 40  # physical activity never above 40 minutes

# Notes are composed from an opener and a detail, so the pool is large enough
# that no two rows in the sheet ever carry the same remark.
OPENERS_WEEKDAY = [
    "Classes ran back to back",
    "Lecture heavy morning",
    "Lab session took most of the afternoon",
    "Slow start, but the day picked up",
    "Assignment deadline kept me at the desk",
    "Quiet day on campus",
    "Double lab today",
    "Group project meeting after class",
    "Presentation prep ate the evening",
    "Stayed back in the library",
    "Tutorial went longer than planned",
    "Ordinary timetable, nothing unusual",
]
OPENERS_WEEKEND = [
    "No classes today",
    "Slept in and took it easy",
    "Spent the day off campus",
    "Weekend pace, nothing scheduled",
    "Caught up on everything I had pushed back",
    "Lazy start to the day",
    "Stayed in for most of it",
    "Ran errands through the afternoon",
    "Long call with family in the evening",
    "Sat with side project work",
]
DETAILS = [
    "kept the coding light",
    "managed a short walk after",
    "caught up on notes at night",
    "skipped the workout for once",
    "ended up reading past midnight",
    "the canteen queue was endless",
    "rain on the way back",
    "internet was patchy all evening",
    "felt sharp right through the afternoon",
    "energy dipped badly after lunch",
    "finished the day ahead of schedule",
    "went to bed later than I meant to",
]


def read_history(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if SHEET not in wb.sheetnames:
        sys.exit(f"error: {path} has no {SHEET!r} sheet")
    ws = wb[SHEET]
    rows = []
    for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
        date = ws.cell(row, COLUMNS["date"]).value
        if not isinstance(date, (dt.datetime, dt.date)):
            continue
        entry = {"date": date if isinstance(date, dt.datetime) else dt.datetime(date.year, date.month, date.day)}
        for key, col in COLUMNS.items():
            if key == "date":
                continue
            entry[key] = ws.cell(row, col).value
        rows.append(entry)
    if not rows:
        sys.exit(f"error: no dated rows found in {path}")
    rows.sort(key=lambda r: r["date"])
    return rows


def sample_duration(rng, values):
    """Pick a value near the recent median, clamped to the observed range."""
    values = [v for v in values if isinstance(v, (int, float))]
    if not values:
        return 0
    median = statistics.median(values)
    low, high = min(values), max(values)
    spread = max(5.0, 0.2 * median)
    value = int(round(rng.gauss(median, spread)))
    return max(int(low), min(int(high), max(0, value)))


def weighted_choice(rng, values):
    values = [v for v in values if v]
    if not values:
        return None
    options = sorted(set(values))
    weights = [values.count(v) for v in options]
    return rng.choices(options, weights=weights, k=1)[0]


def pick_note(rng, rows, date, is_free_day):
    """Return a remark no other row in the sheet already uses."""
    used = {
        str(r["notes"]).strip().lower()
        for r in rows
        if r.get("notes") not in (None, "")
    }
    openers = OPENERS_WEEKEND if is_free_day else OPENERS_WEEKDAY
    candidates = [f"{o}, {d}" for o in openers for d in DETAILS]
    rng.shuffle(candidates)
    for note in candidates:
        if note.strip().lower() not in used:
            return note
    # Pool exhausted (a year-plus of entries): date-stamp keeps it unique.
    return f"{candidates[0]} ({date.strftime('%d %b')})"


def generate(rows, date):
    rng = random.Random(f"daily-log:{date.isoformat()}")
    recent = [r for r in rows if (date - r["date"].date()).days <= HISTORY_DAYS] or rows[-HISTORY_DAYS:]

    entry = {"date": date.isoformat()}
    for key in DURATIONS:
        entry[key] = sample_duration(rng, [r[key] for r in recent])
    entry["sleep"] = min(entry["sleep"], MAX_SLEEP)
    entry["fitness"] = min(entry["fitness"], MAX_FITNESS)

    same_weekday = [r for r in rows if r["date"].weekday() == date.weekday()][-4:]
    pairs = [
        (int(r["class_min"] or 0), int(r["classes"] or 0))
        for r in same_weekday
        if r["class_min"] is not None
    ]
    class_min, classes = rng.choice(pairs) if pairs else (0, 0)
    if classes > MAX_CLASSES:
        class_min = min(class_min, MAX_CLASSES * 60)
        classes = MAX_CLASSES
    entry["class_min"], entry["classes"] = class_min, classes

    entry["feeling"] = weighted_choice(rng, [r["feeling"] for r in recent]) or "Good"
    entry["satisfaction"] = weighted_choice(rng, [r["satisfaction"] for r in recent]) or "Satisfied"
    entry["energy"] = weighted_choice(rng, [r["energy"] for r in recent]) or "Medium"

    entry["notes"] = pick_note(rng, rows, date, entry["class_min"] == 0)

    # Keep the tracked total inside a real day.
    while sum(entry[k] for k in DURATIONS + ["class_min"]) > 1380:
        entry["other"] = max(0, entry["other"] - 15)
        if entry["other"] == 0:
            entry["sleep"] = max(300, entry["sleep"] - 15)
            if entry["sleep"] == 300:
                break
    return entry


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--workbook",
        default=str(Path(__file__).resolve().parent.parent / "mca31.xlsx"),
        help="workbook to read history from (default: mca31.xlsx at the repo root)",
    )
    ap.add_argument("--date", help="YYYY-MM-DD (default: today)")
    ap.add_argument("--out", help="write the JSON here instead of stdout")
    args = ap.parse_args()

    date = (
        dt.datetime.strptime(args.date, "%Y-%m-%d").date()
        if args.date
        else dt.date.today()
    )
    entry = generate(read_history(Path(args.workbook)), date)
    text = json.dumps(entry, indent=2) + "\n"
    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text)
        print(f"wrote {out}")
    else:
        sys.stdout.write(text)


if __name__ == "__main__":
    main()
