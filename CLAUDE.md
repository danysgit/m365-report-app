# M365 Tenant Discovery Report App

## What this app does
Self-hosted Flask web app that takes Fly Migration tenant discovery Excel reports (.xlsx) and generates interactive dashboards with PDF export.

## Supported workloads
- Exchange Online (mailbox details, storage, types, archive status)
- Microsoft 365 Groups (privacy, owners, members, guests)
- Microsoft Teams (channels, privacy, storage)
- OneDrive (user storage, files, activity)
- SharePoint Online (site collections, templates, files)

## Architecture
- **Backend**: Flask (app.py) — parsers, stats computation, PDF generation with matplotlib
- **Frontend**: Jinja2 templates with Chart.js for interactive charts
- **Data flow**: Upload .xlsx → parse with openpyxl → compute stats → store as JSON file in uploads/ → render dashboard → export PDF with matplotlib
- **Session**: Stats stored as JSON files (not cookies — too large at ~30KB)

## Key files
- `app.py` — All backend logic (parsers, stats, PDF generation, routes)
- `templates/index.html` — Multi-file upload page with drag & drop
- `templates/dashboard.html` — Dashboard with Chart.js charts and sortable tables

## PDF export
Two modes: Summary (charts only, 1 page per workload) and Detailed (charts + data tables)

## Tech decisions
- File-based stats storage instead of Flask session cookies (data exceeds 4KB cookie limit)
- Separate Browse buttons per upload slot to avoid double file-dialog bug
- Chart.js for browser charts, matplotlib for PDF charts
- Flexible column parsing with openpyxl to handle variations in Fly report format