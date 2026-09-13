"""
embed_dashboards.py

Turns CEO_Dashboard.html (currently a thin shell that loads
CEO_Dashboard_Standalone.html / CEO_Dashboard_Conso.html via iframe `src=`
file references) into a single, fully self-contained file -- suitable for
hosting alone, with no sibling files needed on the server.

How: base64-encodes the full content of both dashboard files and embeds
them as `data:text/html;base64,...` URIs directly in the shell's iframe
`src` attributes, replacing the external file references. This keeps the
iframe-based isolation (each dashboard still runs in its own JS realm --
the whole reason iframes were used in the first place, since both
dashboards declare colliding global variable/function names) while making
the outer file independent of any other file on disk.

Re-run this any time CEO_Dashboard_Standalone.html or CEO_Dashboard_Conso.html
are rebuilt (refresh_ceo_dashboard.py / build_ceo_dashboard_conso.py), since
the embedded copies are a snapshot taken at embed time, not a live reference.
"""

import base64
import re
from pathlib import Path

BASE = Path(__file__).resolve().parent
SHELL_PATH = BASE / "CEO_Dashboard.html"
STANDALONE_PATH = BASE / "CEO_Dashboard_Standalone.html"
CONSO_PATH = BASE / "CEO_Dashboard_Conso.html"


def to_data_uri(path):
    data = path.read_bytes()
    b64 = base64.b64encode(data).decode("ascii")
    return f"data:text/html;charset=utf-8;base64,{b64}"


def main():
    shell = SHELL_PATH.read_text(encoding="utf-8")

    standalone_uri = to_data_uri(STANDALONE_PATH)
    conso_uri = to_data_uri(CONSO_PATH)

    new_shell, n1 = re.subn(
        r'(<iframe id="frame-standalone"[^>]*\bsrc=")[^"]*(")',
        lambda m: m.group(1) + standalone_uri + m.group(2),
        shell,
    )
    new_shell, n2 = re.subn(
        r'(<iframe id="frame-conso"[^>]*\bsrc=")[^"]*(")',
        lambda m: m.group(1) + conso_uri + m.group(2),
        new_shell,
    )
    if n1 != 1 or n2 != 1:
        raise RuntimeError(f"Expected exactly 1 replacement each for frame-standalone/frame-conso, got {n1}/{n2}")

    SHELL_PATH.write_text(new_shell, encoding="utf-8")
    total_kb = len(new_shell.encode("utf-8")) / 1024
    print(f"Embedded Standalone ({STANDALONE_PATH.stat().st_size/1024:.0f} KB) "
          f"and Conso ({CONSO_PATH.stat().st_size/1024:.0f} KB) into {SHELL_PATH.name}. "
          f"New file size: {total_kb:.0f} KB.")


if __name__ == "__main__":
    main()
