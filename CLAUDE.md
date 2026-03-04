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
- `static/favicon.svg` — SVG favicon used in browser tabs (blue-to-purple gradient bar chart)
- `static/favicon.png` — PNG version of favicon used as Unraid container icon
- `Dockerfile` — Production image using gunicorn (2 workers), includes Unraid docker labels
- `unraid-template.xml` — Unraid Community Apps template with port/path/secret key config
- `docker-compose.yml` — Local testing (host port 7880 → container port 5000)
- `.github/workflows/docker.yml` — Auto-builds and pushes image to ghcr.io on every push to main

## PDF export
Two modes: Summary (charts only, 1 page per workload) and Detailed (charts + data tables)

## Tech decisions
- File-based stats storage instead of Flask session cookies (data exceeds 4KB cookie limit)
- Separate Browse buttons per upload slot to avoid double file-dialog bug
- Chart.js for browser charts, matplotlib for PDF charts
- Flexible column parsing with openpyxl to handle variations in Fly report format
- gunicorn instead of Flask dev server in Docker (production-ready)
- PNG favicon required for Unraid icon (SVG not supported by Unraid UI)

## Docker / Unraid deployment
- Image published to `ghcr.io/danysgit/m365-report-app:latest` via GitHub Actions on every push to main
- Unraid template URL: `https://raw.githubusercontent.com/danysgit/m365-report-app/main/unraid-template.xml`
- Default host port: **7880** (avoids common Unraid conflicts)
- Persistent storage: `/app/uploads` mapped to `/mnt/user/appdata/m365-report/uploads`
- `SECRET_KEY` env var for Flask session security (generate with `openssl rand -hex 32`)
- Unraid docker labels: `net.unraid.docker.webui`, `net.unraid.docker.icon`, `net.unraid.docker.managed`
- Repo is public on GitHub so Unraid can fetch the template XML and icon without auth
