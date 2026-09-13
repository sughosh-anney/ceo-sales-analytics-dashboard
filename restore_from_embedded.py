"""
restore_from_embedded.py

CEO_Dashboard_Standalone.html and CEO_Dashboard_Conso.html were deleted
from disk after embed_dashboards.py baked their full content into
CEO_Dashboard.html as base64 data URIs (for standalone VM hosting). Both
build scripts (refresh_ceo_dashboard.py, build_ceo_dashboard_conso.py)
need the two separate files to exist as templates to splice fresh data
into -- this losslessly decodes them back out of CEO_Dashboard.html.

Run this once to restore the two files, then re-run the normal refresh/
build scripts as usual, then re-run embed_dashboards.py at the end to bake
the updated content back into CEO_Dashboard.html for hosting.
"""

import base64
import re
from pathlib import Path

BASE = Path(__file__).resolve().parent
SHELL_PATH = BASE / "CEO_Dashboard.html"
STANDALONE_PATH = BASE / "CEO_Dashboard_Standalone.html"
CONSO_PATH = BASE / "CEO_Dashboard_Conso.html"


def extract(shell, iframe_id):
    m = re.search(rf'<iframe id="{iframe_id}"[^>]*\bsrc="data:text/html;charset=utf-8;base64,([^"]+)"', shell)
    if not m:
        raise RuntimeError(f"Could not find embedded data URI for {iframe_id}")
    return base64.b64decode(m.group(1))


def main():
    shell = SHELL_PATH.read_text(encoding="utf-8")

    if STANDALONE_PATH.exists():
        print(f"{STANDALONE_PATH.name} already exists -- skipping restore.")
    else:
        STANDALONE_PATH.write_bytes(extract(shell, "frame-standalone"))
        print(f"Restored {STANDALONE_PATH.name} ({STANDALONE_PATH.stat().st_size/1024:.0f} KB)")

    if CONSO_PATH.exists():
        print(f"{CONSO_PATH.name} already exists -- skipping restore.")
    else:
        CONSO_PATH.write_bytes(extract(shell, "frame-conso"))
        print(f"Restored {CONSO_PATH.name} ({CONSO_PATH.stat().st_size/1024:.0f} KB)")


if __name__ == "__main__":
    main()
