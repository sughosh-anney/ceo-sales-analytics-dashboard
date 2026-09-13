# Deployment Guide — Internal Static Host (Waitress)

This describes the deployment pattern generically. Every host name, path, port and
network range below is a **placeholder** — substitute your own:

| Placeholder | Meaning |
|---|---|
| `<APP_ROOT>` | Folder the dashboard files are deployed into on the host |
| `<APP_HOST>` | Host name or address people browse to |
| `<PORT>` | TCP port the server listens on |
| `<INTERNAL_SUBNET>` | The internal/VPN client address range allowed to reach it |

The build produces three self-contained static HTML files — data baked in, no
database, no API calls: `CEO_Dashboard.html` (the merged shell with the
Standalone/Consolidated switcher — this is the one to link people to), plus
`CEO_Dashboard_Standalone.html` and `CEO_Dashboard_Conso.html` (embedded inside
the shell via iframe, but also directly browsable on their own).

Rather than a full web server, this is served via **Waitress** — a
production-quality, pure-Python WSGI server. No web server install, no admin
rights, works identically on any machine with Python.

## Prerequisites on the host

- Python 3.x installed.
- Install dependencies:
  ```
  pip install waitress openpyxl
  ```
- Copy the dashboard folder to `<APP_ROOT>` — needs `CEO_Dashboard.html`,
  `CEO_Dashboard_Standalone.html`, `CEO_Dashboard_Conso.html`,
  `serve_ceo_dashboard.py`, `refresh_ceo_dashboard.py`,
  `build_ceo_dashboard_conso.py`, and the current source workbooks (needed for
  refreshing data, not for serving).
- Set the source-workbook environment variables the refresh scripts read
  (`CONSO_MIS_WORKBOOK`, `CONSO_MIS_WORKBOOK_AUG`, `CONSO_MIS_WORKBOOK_JUL`, and
  any others named in the script headers). No source path is hard-coded.

## 1. Serve the dashboard via Waitress

```
cd "<APP_ROOT>"
python serve_ceo_dashboard.py --host 0.0.0.0 --port <PORT>
```

That's it — no site to configure, no default-document setting, no app pool.
`serve_ceo_dashboard.py` is a small static-file WSGI app that:

- Serves `CEO_Dashboard.html` by default at `http://<APP_HOST>:<PORT>/`.
- Also serves `CEO_Dashboard_Standalone.html` and `CEO_Dashboard_Conso.html`
  directly if someone bookmarks those URLs.
- Refuses to serve anything outside the folder (blocks `../` path traversal) and
  refuses any file extension other than html/css/js/images — so the source
  workbooks sitting in the same folder can never be fetched over HTTP even if
  someone guesses the filename.
- Sends `Cache-Control: no-cache` on every response, so refreshed dashboard data
  is never masked by a stale browser cache.

Confirm inbound TCP on `<PORT>` is allowed through the host firewall:

```powershell
New-NetFirewallRule -DisplayName "Sales Dashboard" -Direction Inbound `
  -LocalPort <PORT> -Protocol TCP -Action Allow
```

### Keep it running after log-off / reboot

Run interactively and the server stops when the session ends. For a long-lived
host, run it as a **Scheduled Task** instead:

1. Task Scheduler → Create Task.
2. **General**: check "Run whether user is logged on or not".
3. **Triggers**: New → "At startup" (so it survives a reboot).
4. **Actions**: New → "Start a program", Program/script = full path to
   `python.exe`, Arguments = `serve_ceo_dashboard.py --host 0.0.0.0 --port <PORT>`,
   Start in = `<APP_ROOT>`.
5. Run the task once manually to confirm it starts cleanly, then check
   `http://<APP_HOST>:<PORT>/` loads from another machine.

(If a real Windows Service is preferred, Waitress has no built-in service
wrapper, but a service-manager wrapper such as `nssm` can wrap the same command.
Only worth the extra setup if service-manager start/stop/restart controls are
needed.)

## 2. Restrict to internal/VPN-only access

The dashboard carries confidential financial data, so access control matters:

- **Primary control (network level, owned by the network team):** the host should
  sit on a segment reachable only from the corporate network or over VPN — not
  exposed to the public internet. Confirm this with whoever manages the
  VPN/firewall.
- **Defense in depth (host firewall):** rather than allowing the port from any
  source, scope the rule to the internal client range:
  ```powershell
  New-NetFirewallRule -DisplayName "Sales Dashboard" -Direction Inbound `
    -LocalPort <PORT> -Protocol TCP -RemoteAddress <INTERNAL_SUBNET> -Action Allow
  ```
- Waitress has no built-in TLS. If internal traffic must also be HTTPS per
  policy, put a lightweight TLS-terminating reverse proxy in front of Waitress
  bound to loopback, rather than exposing the plain-HTTP port directly. Optional
  if the VPN tunnel is already encrypted end to end.

## 3. Automate the workbook → dashboard refresh

Refreshing Standalone: drop the new workbook cut into `<APP_ROOT>`, run
`python refresh_ceo_dashboard.py`. Refreshing Consolidated: put the new source
cuts in place, point the environment variables at them, run
`python build_ceo_dashboard_conso.py`. To automate:

1. Decide where new cuts land on the host — same `<APP_ROOT>`, via whatever route
   delivers them (shared drive, sync client, manual copy).
2. Create a **Scheduled Task** per script:
   - **Trigger**: daily at a fixed time, or more often if cuts arrive more
     frequently.
   - **Action**: run `python.exe` with argument `refresh_ceo_dashboard.py` (or
     `build_ceo_dashboard_conso.py`), **Start in** = `<APP_ROOT>`.
3. `refresh_ceo_dashboard.py` auto-picks whichever workbook in the folder has the
   newest modified date (warning if more than one exists) and regenerates
   `CEO_Dashboard_Standalone.html` in place; `build_ceo_dashboard_conso.py` reads
   its two configured source locations and regenerates `CEO_Dashboard_Conso.html`.
   Waitress just serves whatever is currently on disk — nothing needs restarting;
   the next browser refresh shows the updated numbers.
4. Archive old cuts into a subfolder periodically so the "most recent file" logic
   never picks up something stale by accident.

## 4. Verification checklist

- [ ] Python + `waitress` + `openpyxl` installed on the host
- [ ] Folder copied to `<APP_ROOT>`, `serve_ceo_dashboard.py` running
      (interactively or as a Scheduled Task with an "At startup" trigger)
- [ ] Source-workbook environment variables set; scripts run without a path error
- [ ] `http://<APP_HOST>:<PORT>/` loads `CEO_Dashboard.html` and both
      Standalone/Consolidated views work inside it
- [ ] Dashboard loads correctly from an internal/VPN-connected client
- [ ] Dashboard does **not** load from outside that network (confirms isolation
      actually works)
- [ ] Source workbooks are NOT fetchable over HTTP
      (`curl http://<APP_HOST>:<PORT>/<filename>.xlsm` returns 404)
- [ ] Refresh scheduled tasks created and run once manually to confirm they work
- [ ] After placing a new cut and running the refresh task, the dashboard's
      numbers update on the next browser refresh (no server restart needed)
