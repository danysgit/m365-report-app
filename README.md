# Microsoft 365 Tenant Discovery Report App

A self-hosted web app that takes Fly Migration tenant discovery reports (.xlsx) and generates interactive dashboards with PDF export.

## Supported Workloads

| Workload | Report Type | Charts Generated |
|---|---|---|
| **Exchange Online** | Mailbox details | Mailboxes by type, storage buckets, archive status, top 10 by storage/items |
| **Microsoft 365 Groups** | Group details | Privacy distribution, SharePoint storage, group details |
| **Microsoft Teams** | Team details | Privacy distribution, channel types, storage |
| **OneDrive** | Site details | Storage distribution, top 10 by file count |
| **SharePoint Online** | Site collection details | Storage buckets, activity timeline, templates, top 10 by files |

## Quick Start

```bash
pip install -r requirements.txt
python app.py
# Open http://localhost:5000
```

## Docker

```bash
docker build -t m365-report .
docker run -p 5000:5000 -e SECRET_KEY=$(python3 -c "import secrets;print(secrets.token_hex(32))") m365-report
```

## Usage

1. Open `http://localhost:5000`
2. Upload one or more Fly Migration `.xlsx` files (any combination)
3. View the interactive dashboard — switch between workloads using the nav bar
4. Each workload has an overview tab (charts) and a details tab (sortable table)
5. **Export to PDF** with two options:
   - **Summary** — charts overview only (1 page per workload)
   - **Detailed** — charts + full data tables (2 pages per workload)

## Project Structure

```
m365-report-app/
├── app.py              # Flask app (parsers, stats, PDF generation, routes)
├── requirements.txt
├── Dockerfile
├── templates/
│   ├── index.html      # Multi-file upload page
│   └── dashboard.html  # Dashboard with Chart.js
└── uploads/            # Uploaded files (auto-created)
```
