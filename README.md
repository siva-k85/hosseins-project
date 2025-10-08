# eMMA Opportunities Automation Suite

A production-ready toolchain for collecting Maryland eMMA procurement opportunities, enriching them with detail-page metadata, and delivering curated Excel workbooks and a Streamlit dashboard for downstream analysis.

## Table of Contents
1. [Project Highlights](#project-highlights)
2. [System Architecture](#system-architecture)
3. [Quick Start](#quick-start)
4. [Running the Scraper](#running-the-scraper)
5. [Streamlit Experience](#streamlit-experience)
6. [Testing & Quality](#testing--quality)
7. [Key Files & Layout](#key-files--layout)
8. [Troubleshooting](#troubleshooting)
9. [Releasing & Git Workflow](#releasing--git-workflow)

## Project Highlights
- **Robust web scraping** – walks the ASP.NET browse pages, handles hidden form fields, avoids pagination loops, and throttles requests adaptively when the site pushes back (403/429).
- **Detail enrichment** – for each opportunity, fetches the detail page and extracts solicitation IDs, summaries, procurement contacts, contact email, instructions, program goals, and due dates.
- **Excel-first pipeline** – writes to a structured workbook (`Master`, `Log`, `Archive`, `Refs`), maintains timestamped backups, and automatically upgrades headers to the 20-column schema.
- **Duplicate & schema resilience** – header alias mapping tolerates column reordering; composite-key deduplication prevents duplicate records even when solicitation IDs shift.
- **CLI ergonomics** – single entrypoint (`main-code.py`) with flags for skipping details, choosing historical days, and adjusting logging; environment variables provide defaults.
- **Streamlit dashboard** – a polished UI (`streamlit_app/app.py`) for filtering opportunities, previewing sheets, visualising run history, and downloading filtered Excel views.
- **Test coverage** – unit tests for date parsing, header extraction, deduplication, and record IDs, using real HTML fixtures.

## System Architecture

```mermaid
flowchart TD
    A[Start Scheduled Run] --> B[Load config & CLI flags]
    B --> C[Configure logging & HTTP session]
    C --> D[Fetch listing page]
    D --> E[Parse table headers & rows]
    E --> F{More pages?}
    F -- yes --> D
    F -- no --> G{fetch_details?}
    G -- yes --> H[Fetch detail pages<br/>progress every 10]
    H --> I[Merge detail data<br/>normalize & deduplicate]
    G -- no --> I
    I --> J[Filter target date (DAYS_AGO)]
    J --> K[Load workbook<br/>ensure headers]
    K --> L[Write Master/Log/Archive rows<br/>update statuses]
    L --> M[Generate analytics & formatting]
    M --> N[Create backup<br/>timestamped copy]
    N --> O[Save workbook]
    O --> P[Log summary & exit]
```

## Quick Start

### 1. Clone & install dependencies
```bash
git clone https://github.com/<your-org>/<repo>.git
cd hossein-project
python -m venv .venv
source .venv/bin/activate  # .venv\Scripts\activate on Windows
pip install -r requirements.txt
```

### 2. Optional: developer extras
```bash
pip install -r requirements-dev.txt
pre-commit install
```

### 3. Configure workbook location (defaults provided)
```bash
export EMMA_XLSX="/Users/<you>/Documents/emma_opportunities.xlsx"  # adjust for your OS
```

## Running the Scraper

The primary entry point is `main-code.py`.

```bash
# Pull today’s records with detail enrichment
python main-code.py

# Scrape yesterday’s listings only
python main-code.py --days-ago 1

# Fast pass without detail pages
python main-code.py --skip-details

# Verbose diagnostics
python main-code.py --log-level DEBUG
```

### Environment Variables

| Variable | Default | Purpose |
|----------|---------|---------|
| `EMMA_XLSX` | platform-specific path in Documents | Workbook output target |
| `DAYS_AGO` | `0` | Day offset to capture (0=today, 1=yesterday) |
| `STALE_AFTER_D` | `7` | Archive rows not seen for N days |
| `MAX_PAGES` | `50` | Pagination limit for listing browse |
| `SLEEP_BETWEEN` | `1.0` | Initial delay between HTTP requests |
| `LOG_LEVEL` | `INFO` | Base log level (DEBUG/INFO/…) |
| `TIMEOUT_SECONDS` | `30` | HTTP timeout per request |
| `USER_AGENT` | `Mozilla/5.0 (compatible; MD-EmmaScraper/1.0)` | Override UA string |

### Output
- `Master` – current opportunities, one row per record, status field highlights changes.
- `Log` – append-only audit trail for each run and action (New/Updated/Stale/Unchanged).
- `Archive` – pruned rows older than `STALE_AFTER_D` days.
- `Refs` – freeform sheet for lookup/tagging rules.
- `/backups` – timestamped `.xlsx` backups created before every save.

## Streamlit Experience

Launch the dashboard to explore the workbook interactively:

```bash
cd streamlit_app
pip install -r requirements.txt   # optional if already installed
streamlit run app.py
```

Features:
- Sidebar path selector for alternate workbooks.
- Summary metrics (active, new, updated, due soon).
- Search/filter with optional due date slider (Master).
- Sheet selector (Master/Log/Archive/Refs) with download buttons for filtered data.
- Area chart of recent run activity (from the Log sheet).

## Testing & Quality

Pytest covers critical helpers:

```bash
pytest
```

Key tests include:
- Multiple timestamp formats (`tests/test_date_parsing.py`).
- Header alias extraction and duplicate handling (`tests/test_scraping.py`).
- HTML fixture in `tests/fixtures/emma_reordered_columns.html` ensures resilience to column order changes.

## Key Files & Layout

```
├── main-code.py                 # primary scraper entrypoint
├── streamlit_app/
│   ├── app.py                   # Streamlit dashboard
│   └── requirements.txt         # UI dependencies
├── documentations/
│   └── main-code.md             # Deep-dive technical notes
├── tests/
│   ├── conftest.py              # dynamic loader for main-code module
│   ├── test_date_parsing.py
│   ├── test_scraping.py
│   └── fixtures/
│       └── emma_reordered_columns.html
└── requirements*.txt            # runtime/dev dependency manifests
```

## Troubleshooting

| Symptom | Cause | Resolution |
|---------|-------|------------|
| `Failed to fetch detail page …` | transient network issue or throttling | Automatically retried with adaptive backoff; re-run if persistent |
| `Workbook not found` | Incorrect `EMMA_XLSX` path | Set environment variable or provide full path in CLI |
| No rows scraped | Selected day has no listings | Try a different `--days-ago` value |
| Excel locked during save | Workbook open in Excel | Close the file; rerun to resume |

## Releasing & Git Workflow

1. Ensure tests pass (`pytest`).
2. Review changes (`git status`, `git diff`).
3. Commit with a descriptive message:
   ```bash
   git add README.md streamlit_app tests main-code.py documentations/main-code.md
   git commit -m "Document Streamlit UI and add dashboard"
   ```
4. Push to GitHub:
   ```bash
   git push origin <branch-name>
   ```
5. Open a pull request summarising scraper changes, test results, and UI improvements.

> **Note:** This environment cannot perform the push on your behalf. Run the above commands locally (or via your CI/CD runner) using credentials authorised for the GitHub repository.

Happy scraping! 🎯
