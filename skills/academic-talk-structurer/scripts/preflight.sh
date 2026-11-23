#!/bin/sh
# Combined S4 gate: layout+notes lint, number traceability, evidence locators, page budget.
# Usage: preflight.sh <out.pptx> <workdir> <case.json> [spec] [source_txt]
# Optional inputs auto-detected in workdir: number_allow.json, budget_justify.json, figs/manifest.json.
# Exit: 0 all clean (justified entries allowed), 1 any hard issue. Individual commands: references/checks.md
set -u
PPTX="${1:?pptx path}"; WORK="${2:?workdir}"; CASE="${3:?case.json}"
SPEC="${4:-$WORK/deck_spec.json}"; SOURCE="${5:-$WORK/paper.txt}"
SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
# pick a python that can import pptx (PYTHON env wins; then PATH candidates)
pick_py() {
  for c in "${PYTHON:-}" python3 python3.13 python3.12 python3.11 python; do
    [ -z "$c" ] && continue
    command -v "$c" >/dev/null 2>&1 || continue
    if "$c" -c "import pptx" >/dev/null 2>&1; then printf '%s' "$c"; return 0; fi
  done
  printf '%s' "${PYTHON:-python3}"
}
PY=$(pick_py)
"$PY" -c "import pptx" >/dev/null 2>&1 || echo "warn: no interpreter with python-pptx found; set PYTHON=<path>. layout lint will be skipped." 
echo "python: $PY"
fail=0
echo "== preflight gate =="
echo "-- [1/4] layout + notes/timing lint"
if "$PY" -c "import pptx" >/dev/null 2>&1; then
  "$PY" "$SCRIPT_DIR/check_layout.py" "$PPTX" --json "$WORK/lint_report.json" --plan "$WORK/slide_plan.json" || fail=1
else
  echo "   skip (python-pptx unavailable; do S5 visual check instead)"
fi
echo "-- [2/4] number traceability"
if [ -f "$SPEC" ] && [ -f "$SOURCE" ]; then
  ALLOW=""; [ -f "$WORK/number_allow.json" ] && ALLOW="--allow $WORK/number_allow.json"
  # shellcheck disable=SC2086
  "$PY" "$SCRIPT_DIR/check_numbers.py" --spec "$SPEC" --source "$SOURCE" --out "$WORK/number_check.json" $ALLOW || fail=1
else
  echo "   skip (missing spec or source text)"
fi
echo "-- [3/4] evidence locator existence"
if [ -f "$SPEC" ] && [ -f "$SOURCE" ]; then
  MAN=""; PGS=""
  [ -f "$WORK/figs/manifest.json" ] && MAN="--manifest $WORK/figs/manifest.json"
  [ -f "$WORK/inspect.json" ] && PGS=$( [ -f "$SOURCE" ] && "$PY" -c "
import json,sys
try:
    d=json.load(open('$WORK/inspect.json'))
    sp=d.get('source_pdf') or {}
    print('--pages', sp.get('pages')) if sp.get('pages') else None
except Exception: pass" )
  # shellcheck disable=SC2086
  "$PY" "$SCRIPT_DIR/check_evidence.py" --spec "$SPEC" --source "$SOURCE" $MAN $PGS --out "$WORK/evidence_check.json" || fail=1
else
  echo "   skip (missing spec or source text)"
fi
echo "-- [3b] spec figure paths exist"
if [ -f "$SPEC" ]; then
  "$PY" - "$SPEC" "$WORK" <<'QEOF'
import json, sys
from pathlib import Path
spec, work = sys.argv[1], sys.argv[2]
bad = []
for i, s in enumerate(json.load(open(spec)).get("slides", []), 1):
    f = s.get("figure") or {}
    p = f.get("path")
    if not p:
        continue
    if not Path(p).exists():
        # relative-to-workdir fallback (extract_figs canonical layout)
        alt = [Path(work) / "figs" / Path(p).name, Path(work) / Path(p).name]
        if not any(a.exists() for a in alt):
            bad.append((i, p))
for i, p in bad:
    print(f"  ! [s{i}] figure path missing: {p}")
sys.exit(1 if bad else 0)
QEOF
  [ $? -eq 0 ] || fail=1
fi
echo "-- [4/4] page budget vs audience"
if [ -f "$SPEC" ] && [ -f "$CASE" ]; then
  J=""; [ -f "$WORK/budget_justify.json" ] && J="--justif $WORK/budget_justify.json"
  # shellcheck disable=SC2086
  "$PY" "$SCRIPT_DIR/plan_pages.py" check --case "$CASE" --spec "$SPEC" $J --out "$WORK/budget_check.json" || fail=1
else
  echo "   skip (missing spec or case)"
fi
echo "== preflight done: exit $fail =="
exit $fail
