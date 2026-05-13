# Salesforce Critical & Urgent Case Monitor

> A Flask + SOQL dashboard that surfaces open Critical/Urgent Salesforce cases, flags them against SLA windows, exports a styled Excel report, and (optionally) hands the report off to SharePoint.

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Flask-3.x-000000.svg?logo=flask)](https://flask.palletsprojects.com/)
[![Salesforce](https://img.shields.io/badge/Salesforce-API-00A1E0.svg?logo=salesforce&logoColor=white)](https://developer.salesforce.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## Why this project

Support teams live and die by SLA. When a Critical or Urgent ticket sits open too long, customers escalate and CSAT tanks. This project gives the on-call lead a single page to answer:

> *"Which Critical/Urgent cases are open, who owns them, and how close are we to breaching SLA?"*

Cases are pulled live from Salesforce via SOQL, enriched with owner contact info, classified against three SLA age windows (12h / 30h / 45h), rendered into a searchable web dashboard, and exported to a stakeholder-ready Excel report — with an optional PowerShell handoff to publish that report to SharePoint.

## Features

- **Live SOQL query** — pulls open `Critical` / `Urgent` cases created within a configurable lookback window (default: 45h), excluding `Closed` and `Closed Pending` statuses.
- **Owner enrichment** — second SOQL hop into the `User` object to attach owner name and email; results are cached per request to keep API call counts low.
- **SLA flagging** — each case is tagged `Alert_12h`, `Alert_30h`, or `Alert_45h` based on age, surfaced in the UI as red badges so the on-call can triage at a glance.
- **Searchable dashboard** — Flask + Jinja2 page with a single search box that filters across every visible column client-side.
- **Styled Excel export** — pandas + openpyxl produces a `status_report.xlsx` with a real Excel Table (filter, sort, banded rows) ready to drop into a stakeholder email.
- **Optional SharePoint upload** — environment-driven hook to a PowerShell script (PnP.PowerShell) that pushes the report to a document library. Skipped gracefully when not configured or off-Windows.
- **Twelve-factor config** — all secrets and runtime knobs live in `.env`; no credentials are committed.

## Architecture

```
              ┌──────────────────┐
              │  Salesforce Org  │
              │  (Cases + Users) │
              └────────┬─────────┘
                       │ SOQL via simple-salesforce
                       ▼
┌────────────┐   ┌──────────────────────────┐   ┌─────────────────┐
│  Browser   │──▶│  Flask app (app.py)      │──▶│  status_report  │
│  /         │   │  • fetch_open_critical_  │   │  .xlsx (pandas  │
│            │◀──│    cases()               │   │  + openpyxl)    │
│  Dashboard │   │  • SLA age flags         │   └────────┬────────┘
└────────────┘   │  • Jinja2 render         │            │
                 └──────────────────────────┘            ▼
                                              ┌────────────────────┐
                                              │  PowerShell script │
                                              │  (optional, env-   │
                                              │  gated)            │
                                              └─────────┬──────────┘
                                                        ▼
                                              ┌────────────────────┐
                                              │  SharePoint Docs   │
                                              └────────────────────┘
```

## Tech stack

| Layer | Tool |
|---|---|
| Backend | Python 3.10+, Flask 3.x |
| Salesforce | `simple-salesforce` (SOQL over REST) |
| Data | pandas, openpyxl |
| Frontend | Jinja2 templates, vanilla JS, hand-rolled CSS |
| Config | `python-dotenv` (12-factor) |
| Distribution (optional) | PowerShell + PnP.PowerShell → SharePoint |

## Quick start

### 1. Clone

```bash
git clone https://github.com/Nithin8112001/salesforce-case-monitor.git
cd salesforce-case-monitor
```

### 2. Install

```bash
python -m venv .venv
source .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configure

```bash
cp .env.example .env
# then edit .env and fill in your Salesforce credentials
```

You'll need:

- **`SALESFORCE_USERNAME`** / **`SALESFORCE_PASSWORD`** — your Salesforce login.
- **`SALESFORCE_SECURITY_TOKEN`** — generated from *Setup → My Personal Information → Reset My Security Token*.
- **`SALESFORCE_INSTANCE`** — e.g. `https://yourcompany.my.salesforce.com`.

### 4. Run

```bash
python app.py
```

Your default browser opens at `http://127.0.0.1:5000/`. Each page load re-queries Salesforce, regenerates `status_report.xlsx`, and (if configured) re-runs the SharePoint upload script.

## SOQL behind the dashboard

The primary case query (parameterized at runtime by `LOOKBACK_HOURS`):

```sql
SELECT Product_Line__c, Account.Name, CaseNumber, Case_Severity__c,
       OwnerId, CreatedDate
FROM   Case
WHERE  Case_Severity__c IN ('Critical', 'Urgent')
  AND  CreatedDate >= :lookback_iso
  AND  Status NOT IN ('Closed', 'Closed Pending')
ORDER BY CreatedDate DESC
```

Each unique `OwnerId` is then resolved with a cached lookup against the `User` object so the dashboard can show owner name and email without N+1 queries.

## Configuration reference

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `SALESFORCE_USERNAME` | yes | — | SF login |
| `SALESFORCE_PASSWORD` | yes | — | SF password |
| `SALESFORCE_SECURITY_TOKEN` | yes | — | SF security token |
| `SALESFORCE_INSTANCE` | yes | — | SF instance URL |
| `LOOKBACK_HOURS` | no | `45` | How far back to query for new cases |
| `REPORT_FILENAME` | no | `status_report.xlsx` | Excel output filename |
| `FLASK_HOST` | no | `127.0.0.1` | Bind address |
| `FLASK_PORT` | no | `5000` | Bind port |
| `FLASK_DEBUG` | no | `true` | Flask debug mode |
| `AUTO_OPEN_BROWSER` | no | `true` | Pop the browser on startup |
| `POWERSHELL_SCRIPT_PATH` | no | *(unset)* | Absolute path to the SharePoint uploader; leave blank to skip |

## Project structure

```
.
├── app.py                  # Flask app, SOQL, Excel export, SharePoint hook
├── templates/
│   └── index.html          # Dashboard (Jinja2)
├── scripts/
│   └── upload_csv.ps1      # Optional SharePoint uploader (PnP.PowerShell)
├── requirements.txt
├── .env.example            # Copy to .env and fill in secrets
├── .gitignore
├── LICENSE
└── README.md
```

## Roadmap

- Background scheduler (APScheduler) so the dashboard self-refreshes without a page reload.
- Email alerts to case owners when a ticket crosses an SLA threshold.
- Pluggable severity rules (currently fixed at Critical/Urgent + 12/30/45h windows).
- Pytest coverage around `_resolve_owner` and the SLA-flag math.

## License

[MIT](LICENSE) © Nithin Gopal

---

Built by **[Nithin Gopal](https://github.com/Nithin8112001)** while supporting Salesforce-driven case operations. If you're hiring for a Salesforce / CRM developer role, I'd love to chat.
