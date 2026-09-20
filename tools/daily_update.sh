#!/usr/bin/env bash
# One day's update: build the entry, write it into the workbook, commit, push.
#
#   tools/daily_update.sh                 # today
#   tools/daily_update.sh 2026-09-21      # a specific date
#   ENTRY=path/to.json tools/daily_update.sh 2026-09-21   # use given values
#
# Safe to run twice for the same date: the row is overwritten, not duplicated,
# and nothing is committed when the workbook ends up unchanged.
set -euo pipefail

repo="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo"

date_arg="${1:-$(date +%F)}"
workbook="${WORKBOOK:-mca31.xlsx}"

python3 -c "import openpyxl" 2>/dev/null || pip install --quiet openpyxl

entry="${ENTRY:-pending/${date_arg}.json}"
generated=0
if [[ ! -f "$entry" ]]; then
  python3 tools/generate_entry.py --workbook "$workbook" --date "$date_arg" --out "$entry"
  generated=1
fi

# A date already in the sheet is left alone — an existing row is never rewritten
# by the daily job. Use append_daily_row.py directly to change one on purpose.
result="$(python3 tools/append_daily_row.py --workbook "$workbook" --entry "$entry" --if-exists skip)"
echo "$result"
if [[ "$result" == skipped:* ]]; then
  (( generated )) && rm -f "$entry"
  exit 0
fi

git add "$workbook" "$entry"
if git diff --cached --quiet; then
  echo "nothing to commit for ${date_arg}"
  exit 0
fi

git -c user.name="${GIT_AUTHOR_NAME:-Ananye Dwivedi}" \
    -c user.email="${GIT_AUTHOR_EMAIL:-ananye532@gmail.com}" \
    commit -q -m "Daily log: ${date_arg}"

branch="$(git rev-parse --abbrev-ref HEAD)"
for delay in 2 4 8 16 0; do
  if git push -u origin "$branch"; then
    echo "pushed ${branch}"
    exit 0
  fi
  [[ "$delay" == 0 ]] && break
  echo "push failed, retrying in ${delay}s"
  sleep "$delay"
done
echo "push failed after retries" >&2
exit 1
