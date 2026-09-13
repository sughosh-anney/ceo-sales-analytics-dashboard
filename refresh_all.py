"""
refresh_all.py

Since CEO_Dashboard.html is now the ONLY dashboard file kept on disk (its
two source dashboards live embedded inside it as base64 data URIs), a full
data refresh needs 4 steps every time: restore the two source files,
rebuild each one's data, re-embed them, then delete the intermediates
again. This script runs all 4 in sequence.

Usage:
    python refresh_all.py

Prerequisite: drop the new source Excel/xlsm cuts in this folder first,
same as before (refresh_ceo_dashboard.py picks the newest .xlsx by
modified date; build_ceo_dashboard_conso.py reads its configured source locations).
"""

import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
STANDALONE_PATH = BASE / "CEO_Dashboard_Standalone.html"
CONSO_PATH = BASE / "CEO_Dashboard_Conso.html"

STEPS = ["restore_from_embedded.py", "refresh_ceo_dashboard.py", "build_ceo_dashboard_conso.py", "embed_dashboards.py"]


def run(script):
    print(f"\n=== {script} ===")
    result = subprocess.run([sys.executable, str(BASE / script)], cwd=BASE)
    if result.returncode != 0:
        print(f"FAILED at {script} (exit {result.returncode}) -- stopping.")
        sys.exit(result.returncode)


def main():
    for step in STEPS:
        run(step)

    for p in (STANDALONE_PATH, CONSO_PATH):
        if p.exists():
            p.unlink()
            print(f"Removed intermediate {p.name}")

    print("\nDone. CEO_Dashboard.html is up to date and self-contained.")


if __name__ == "__main__":
    main()
