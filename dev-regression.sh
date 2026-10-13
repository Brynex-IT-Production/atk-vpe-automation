#!/usr/bin/env bash
set -euo pipefail

BRANCH="main"
EXPECTED_SCRIPT_ID="1LS3HabXjXOainAMiNLVVlRX5LEaA-o5Y7Kf31Fg-JQJZsR4ZKHI5zHUN"
DEV_DEPLOYMENT_ID="AKfycbwzHdJptr44J1aWN4UqHp6qN435kwFPbG8ngYGEYABqypF2urE0Jyd0hLd63M-wYVVE"
DEV_URL="https://script.google.com/macros/s/${DEV_DEPLOYMENT_ID}/exec"

fail() {
  echo "DEV REGRESSION: FAIL ❌"
  echo "$1"
  exit 1
}

command -v git >/dev/null || fail "git is not installed."
command -v clasp >/dev/null || fail "clasp is not installed."
command -v curl >/dev/null || fail "curl is not installed."
command -v python3 >/dev/null || fail "python3 is not installed."

if [ ! -f ".clasp.json" ] || [ ! -f "Code.js" ] || [ ! -f "WebApiV1.js" ]; then
  fail "Run this command from the atk-vpe-automation repository."
fi

python3 - "$EXPECTED_SCRIPT_ID" <<'PY'
import json, sys
from pathlib import Path
expected = sys.argv[1]
actual = json.loads(Path(".clasp.json").read_text()).get("scriptId", "")
if actual != expected:
    raise SystemExit(f"Refusing deployment: expected DEV script {expected}, found {actual}")
PY

if [ -n "$(git status --porcelain)" ]; then
  fail "Local repository has uncommitted changes. Commit/stash them before DEV deployment."
fi

git pull --ff-only origin "$BRANCH"

if command -v node >/dev/null; then
  while IFS= read -r file; do
    node --check "$file" >/dev/null
  done < <(find . -maxdepth 1 -type f -name '*.js' -print)
fi

clasp push --force
clasp update-deployment "$DEV_DEPLOYMENT_ID" --description "Terminal DEV regression $(git rev-parse --short HEAD)"

verify=$(python3 - <<'PY'
import re
from pathlib import Path
source = Path("Code.js").read_text()
match = re.search(r"const VERIFY_TOKEN\s*=\s*\'([^\']+)\'", source)
if not match:
    raise SystemExit("VERIFY_TOKEN not found in Code.js")
print(match.group(1))
PY
)

python3 - "$verify" <<'PY'
import json, sys
from pathlib import Path
Path("/tmp/atk-vpe-master-payload.json").write_text(json.dumps({
    "system_action": "DEV_MASTER_REGRESSION",
    "verify_token": sys.argv[1]
}))
PY

http_code=$(curl --location --silent --show-error \
  --max-time 330 \
  --output /tmp/atk-vpe-master.json \
  --write-out '%{http_code}' \
  --header 'Content-Type: application/json' \
  --data-binary @/tmp/atk-vpe-master-payload.json \
  "$DEV_URL")

[ "$http_code" = "200" ] || fail "DEV endpoint returned HTTP $http_code."

response=$(cat /tmp/atk-vpe-master.json)

if [ "$response" = "EVENT_RECEIVED" ]; then
  fail "Regression request fell through to the generic webhook handler."
fi

python3 - <<'PY'
import json
from pathlib import Path

raw = Path("/tmp/atk-vpe-master.json").read_text().strip()
try:
    report = json.loads(raw)
except Exception:
    raise SystemExit("DEV REGRESSION: FAIL ❌\nResponse was not JSON:\n" + raw)

if report.get("diagnostic") == "DEV_MASTER_REGRESSION_GUARD":
    raise SystemExit(
        "DEV REGRESSION: FAIL ❌\n"
        "Regression guard rejected the request: "
        f"token_ok={report.get('token_ok')} "
        f"system_mode={report.get('system_mode')!r}"
    )

sections = [
    ("Grammarian parser", "grammarian_parser"),
    ("Planner/API lifecycle", "planner_api_lifecycle"),
    ("Core", "core"),
    ("Configured workflows", "configured_workflows"),
    ("Forced workflows", "forced_enabled_workflows"),
]

print()
print("==============================")
print("DEV REGRESSION SUMMARY")
print("==============================")
print("Version:", report.get("version", "unknown"))
print("Overall:", "PASS ✅" if report.get("ok") is True else "FAIL ❌")

all_good = report.get("ok") is True
for label, key in sections:
    section = report.get(key, {}) or {}
    passed = section.get("checks_passed")
    total = section.get("checks_total")
    ok = section.get("ok") is True and passed == total
    all_good = all_good and ok
    print(f"{label}: {passed}/{total} " + ("✅" if ok else "❌"))

print("==============================")

if not all_good:
    print()
    print(json.dumps(report, indent=2))
    raise SystemExit(1)
PY
