# CEO Sales Analytics Dashboard

> A leadership sales-and-P&L dashboard that ships as **one HTML file**. No database, no build step, no API — the pipeline bakes the data into the page.

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20parsing-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![Chart.js](https://img.shields.io/badge/Chart.js-FF6384?style=flat-square&logo=chartdotjs&logoColor=white)
![Waitress](https://img.shields.io/badge/Waitress-WSGI-306998?style=flat-square)
![Single File](https://img.shields.io/badge/Output-single%20HTML%20file-success?style=flat-square)
![Dependencies](https://img.shields.io/badge/Front--end%20deps-1-success?style=flat-square)

**Built at ADF Foods Ltd — one of 14 dashboards, pipelines and automations delivered between 16 July and 12 September 2026, in under 8 weeks.**

---

## What This Does

Leadership needed one place to see how the group was selling and earning — by entity, zone, category, brand, salesman and SKU — refreshed from the workbooks finance already produced each month, and reachable from any laptop on the internal network with nothing installed.

This repository is the Python pipeline behind that dashboard: **about 2,830 lines across 8 scripts** that read the month's source workbooks and emit a Standalone view and a Consolidated view of all **7 group entities** as a single self-contained HTML file.

The data is baked in as literal JavaScript. That removes a whole category of problems: no database to provision, no API to keep alive, no build toolchain on the host, no version skew between a page and the numbers behind it. Copy the file and you have copied the dashboard.

---

## Key Features

- **Isolated view shells** — the two views are base64-embedded into one shell as separate iframes, so each keeps its own JavaScript realm and identically-named globals cannot collide
- **Recomputed classifications** — growth signal, driver and trend badges are derived in Python from the published threshold rules, so summary and detail tabs can never disagree
- **Reconciled consolidation** — consolidated figures are derived from the transaction rows and reconciled to the management report, entity by entity and month by month
- **Stated-basis reporting** — pre-elimination and elimination-adjusted totals are each used where they belong, and labelled on the page
- **Reversible embedding** — the deployed file decodes losslessly back into its two views, so a refresh starts from the artefact that is live
- **Allow-listed server** — Waitress serves html, css, js, images and fonts only and refuses path traversal, so workbooks in the same folder cannot be fetched over HTTP

---

## Architecture

```
      Source workbooks (locations from the environment)
   sales analytics . group sales data . board MIS report
     annual category sales . financial statements
                            |
             +--------------+--------------+
             |                             |
      Standalone builder            Consolidated builder
      KPIs . profitability          7 entities . entity P&L
      volume . sales . growth       zone . category . brand
             |                             |
             +--------------+--------------+
                            |
                     base64 embed step
                            |
                            v
              one self-contained HTML file
                            |
                            v
               allow-listed static server
                            |
                            v
                  Leadership, on the LAN
```

---

## Tools & Technologies

| Component | Technology |
|---|---|
| Workbook parsing | Python with openpyxl, read-only |
| Aggregation and badge rules | Plain Python, no dataframe layer |
| Charts | Chart.js, the only front-end library |
| Packaging | Base64 embedding into one HTML shell |
| Hosting | Waitress WSGI behind an extension allow-list |
| Configuration | Workbook locations and sheet names from the environment |

---

## Project Context

The numbers move once a month, when finance closes the period, so a rebuilt static artefact is the honest shape for them — restore the views out of the deployed file, rebuild both, re-embed, clean up.

`DEPLOYMENT.md` covers the hosting pattern — scheduled-task startup, firewall scoping to the internal network, a verification checklist — with placeholders for real infrastructure names.

---

*Built at ADF Foods Ltd. The generated dashboard, its source workbooks and all real figures are excluded from this repository.*
