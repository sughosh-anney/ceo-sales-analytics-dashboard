# CEO Sales Analytics Dashboard

> A leadership sales-and-P&L dashboard that ships as **one HTML file**. No database, no build step, no API — the pipeline bakes the data straight into the page.

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20parsing-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![Chart.js](https://img.shields.io/badge/Chart.js-FF6384?style=flat-square&logo=chartdotjs&logoColor=white)
![Waitress](https://img.shields.io/badge/Waitress-WSGI-306998?style=flat-square)
![Single File](https://img.shields.io/badge/Output-single%20HTML%20file-success?style=flat-square)
![Dependencies](https://img.shields.io/badge/Front--end%20deps-1-success?style=flat-square)

**Built at ADF Foods Ltd — one of 14 dashboards, pipelines and automations delivered between 16 July and 12 September 2026, in under 8 weeks.**

---

## What is this project?

Leadership needed one place to see how the group was selling and earning — by entity, zone, category, brand, salesman and SKU — refreshed from the same workbooks finance already produced every month, and reachable from a laptop on the internal network without anyone installing anything.

This is the Python pipeline behind that dashboard: **about 2,800 lines across 8 scripts** that read the source workbooks, recompute everything the source got wrong, and emit **14 tab modules across two views** — a Standalone view and a Consolidated view covering all **7 group entities** — as a single self-contained HTML file.

---

## How the development was done

### One file, on purpose

The output is a single HTML document with its data baked in as literal JavaScript. That decision removes an entire category of problems: no database to provision, no API to keep alive, no build toolchain on the host, no version skew between a page and the data behind it. Copy the file and you have copied the dashboard. Serving it needs nothing more than a static file server.

### Two dashboards, two iframes — and why that is not laziness

The Standalone and Consolidated views are embedded in the shell as **separate iframes**, base64-encoded into the shell's `src` attributes so even the shell is one file.

The reason is concrete: both views declare identically-named globals. Merging their scripts into one document would redeclare the same top-level names and the second one would throw on load. Iframes give each view its own JavaScript realm, so both can keep their natural naming and neither can reach into the other. The alternative — renaming hundreds of symbols across two generated documents on every refresh — would be fragile in exactly the place you least want fragility.

### Recomputing the classifications instead of trusting them

Three badge systems drive how every row is read: a **growth signal**, a **growth driver**, and a **trend** classification. The source workbook carries hand-maintained columns for these, and those columns were inconsistent — categories that had clearly grown were tagged as declining.

So none of that column is read. All three are recomputed in Python from the published threshold rules, including the awkward cases:

- Trend checks trajectory shape *before* the aggregate rules, but a trajectory only qualifies as decelerating if the terminal period's own sign agrees with it. A series that falls month after month while staying positive throughout is steady growth, not deceleration — verified by finding a real row with that exact shape and comparing it against one that does end negative.
- Growth signal, driver and status are derived from the same recomputed numbers everywhere they appear, so the badge on the summary tab and the badge on the detail tab can never disagree.

### Rejecting a source that looked authoritative

The source workbook shipped a pre-built pivot presented as the consolidated total. It was not used — because its "consolidated" figure was tested month by month and found to equal **one single entity's figure in every month**. A filter had been left applied on the pivot and nobody had noticed. Every consolidated number is instead derived from the transaction rows and the board's own MIS report, and each entity-month is reconciled against the P&L before anything downstream is built.

That is the general pattern in this codebase: a figure is used when it has been proved, and the proof is written down in the docstring next to the code that depends on it.

### A file server that cannot leak the source data

The dashboard folder also holds the source workbooks. The bundled Waitress server therefore works from an **extension allow-list** — html, css, js, images and fonts, nothing else — and refuses path traversal. Guessing a workbook's filename over HTTP returns a 404, not a download. The allow-list is a default-deny: adding a new source file type to the folder cannot accidentally expose it.

---

## What was used

Python with openpyxl for reading the source workbooks, Waitress as the WSGI server, and **Chart.js as the only front-end dependency**. No framework, no bundler, no database, no package manager on the client side.

Source workbook locations are read from environment variables, so no machine-specific path is hard-coded anywhere in the repository.

---

## How it works currently

| Script | Role |
|---|---|
| `refresh_ceo_dashboard.py` | Builds the Standalone view — KPI tiles, profitability, volume, sales, top-10 and tail, growth analysis — recomputing every badge and margin from the source cells |
| `build_ceo_dashboard_conso.py` | Builds the Consolidated view for all 7 entities, rescaling each entity-month to the P&L and deriving entity, zone, category and brand breakdowns |
| `update_conso_mis_august.py` | Refreshes the consolidated view onto a later MIS cut as the reporting window widens |
| `update_standalone_entity_pl.py` | Wires the Standalone entity P&L onto live formulas driven by the workbook's own reporting-window selector |
| `embed_dashboards.py` | Base64-embeds both views into the shell, producing the single distributable file |
| `restore_from_embedded.py` | Losslessly decodes both views back out of that file, so a refresh can start from the deployed artefact |
| `refresh_all.py` | Runs restore → rebuild both → re-embed → clean up, as one command |
| `serve_ceo_dashboard.py` | The allow-listed static server |

The working loop is: drop in the new workbook cut, run `refresh_all.py`, done. The server serves whatever is on disk, so nothing needs restarting and the next browser refresh shows the new numbers.

`DEPLOYMENT.md` covers the hosting pattern — scheduled-task startup, firewall scoping to the internal network, and the verification checklist — with placeholders in place of real infrastructure names.

---

## Output

A single self-contained HTML file, served on an internal network to leadership, carrying 14 tab modules across a Standalone and a Consolidated view of all 7 group entities.

Deployment is a file copy and one command:

```
python serve_ceo_dashboard.py --host 0.0.0.0 --port <PORT>
```

---

*Built at ADF Foods Ltd. The generated dashboard, its source workbooks and all real figures are deliberately excluded from this repository.*
