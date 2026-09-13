"""
serve_ceo_dashboard.py

Serves the CEO Dashboard folder's static HTML files via Waitress (a
production-quality pure-Python WSGI server) -- no IIS, no Node, no
external web server needed. All three dashboards are self-contained,
data-baked-in static HTML, so this is a plain static-file server, not an
application server: refreshing the data means re-running
refresh_ceo_dashboard.py / build_ceo_dashboard_conso.py as before, then
this server picks up the new file on the next request automatically
(nothing to restart).

Usage:
    pip install waitress
    python serve_ceo_dashboard.py [--host 0.0.0.0] [--port 8080]

Then browse to:
    http://<host>:8080/                       -> CEO_Dashboard.html (merged shell)
    http://<host>:8080/CEO_Dashboard_Standalone.html
    http://<host>:8080/CEO_Dashboard_Conso.html
"""

import argparse
import mimetypes
from pathlib import Path

from waitress import serve

BASE = Path(__file__).resolve().parent
DEFAULT_DOCUMENT = "CEO_Dashboard.html"

# Only these extensions are ever served -- refuses to serve the source
# .xlsx/.xlsm workbooks or Python scripts even if someone guesses a URL,
# since this folder also holds the (sensitive) data source files.
ALLOWED_EXTENSIONS = {".html", ".htm", ".css", ".js", ".png", ".jpg", ".jpeg",
                       ".gif", ".svg", ".ico", ".woff", ".woff2"}


def app(environ, start_response):
    path = environ.get("PATH_INFO", "/")
    if path == "/":
        path = "/" + DEFAULT_DOCUMENT

    # Resolve against BASE and refuse to leave it (blocks ../.. traversal).
    requested = (BASE / path.lstrip("/")).resolve()
    try:
        requested.relative_to(BASE)
    except ValueError:
        start_response("403 Forbidden", [("Content-Type", "text/plain")])
        return [b"Forbidden"]

    if requested.suffix.lower() not in ALLOWED_EXTENSIONS or not requested.is_file():
        start_response("404 Not Found", [("Content-Type", "text/plain")])
        return [b"Not Found"]

    content_type = mimetypes.guess_type(str(requested))[0] or "application/octet-stream"
    data = requested.read_bytes()
    start_response("200 OK", [
        ("Content-Type", content_type),
        ("Content-Length", str(len(data))),
        ("Cache-Control", "no-cache"),  # dashboards get rebuilt in place; never serve a stale cached copy
    ])
    return [data]


def main():
    parser = argparse.ArgumentParser(description="Serve the CEO Dashboard folder via Waitress.")
    parser.add_argument("--host", default="0.0.0.0", help="Bind address (default 0.0.0.0 -- all interfaces)")
    parser.add_argument("--port", type=int, default=8080, help="Port to listen on (default 8080)")
    args = parser.parse_args()

    print(f"Serving {BASE} via Waitress on http://{args.host}:{args.port}/")
    print(f"Default document: {DEFAULT_DOCUMENT}")
    serve(app, host=args.host, port=args.port)


if __name__ == "__main__":
    main()
