# eMMA Excel Updater

## Purpose
The script `main-code.py` automates retrieval of Maryland eMMA public procurement postings and curates the results inside a single Excel workbook. It keeps an up-to-date view of recent opportunities, preserves an auditable change log, and automatically archives stale items so the workbook does not accumulate outdated rows.

## End-to-End Flow
1. **Load configuration** from environment variables or built-in defaults (line references below). This controls things like workbook location, how many pages to scrape, and logging verbosity.
2. **Initialize logging** with the requested `LOG_LEVEL` so activity is visible in the console.
3. **Establish a resilient HTTP session** (`make_session`, `main-code.py:122-132`) that adds retry/backoff behavior and sets realistic browser headers.
4. **Scrape the eMMA listings** (`emma_scrape`, `main-code.py:242-331`):
   - Request the public browse page, capture hidden ASP.NET form fields, and handle paging via `__doPostBack` events.
   - Parse each row to capture title, URL, category, procurement method, agency, and the publish timestamp.
   - Normalize timestamps to Eastern Time and derive a stable `record_id` using the eMMA numeric ID when available (fallback to a BLAKE2 hash).
   - Filter the staged records down to those published on the target day (`today - DAYS_AGO`).
5. **Prepare the Excel workbook** (`load_wb`, `init_workbook_if_needed`, `main-code.py:349-391`) by creating the required sheets (`Master`, `Log`, `Archive`, `Refs`) when the file is missing or invalid.
6. **Merge staging data into Excel** (`merge_into_excel`, `main-code.py:527-639`):
   - Build an index of existing records by `record_id`.
   - Insert brand-new rows with status `New` and a `first_seen_et` timestamp.
   - For existing rows, detect changes across business fields (title, agency, etc.) and mark them `Updated` or `Unchanged`.
   - Move untouched rows older than `STALE_AFTER_D` days to the `Archive` sheet and tag them `Stale`.
   - Append every action to the `Log` sheet for traceability.
7. **Refresh workbook presentation** by enforcing the `tbl_opps` Excel table, auto-sizing columns, and applying conditional formatting that color-codes status values (`main-code.py:410-515`).
8. **Persist results** to the path defined by `WORKBOOK_PATH` and print a completion message showing the path and target date.

## Key Components
### Configuration & Defaults
- `WORKBOOK_PATH` (`main-code.py:62-65`): Location of the Excel workbook. Defaults to a Windows directory; override with the `EMMA_XLSX` environment variable on other systems.
- `DAYS_AGO`, `STALE_AFTER_D`, `MAX_PAGES`, `SLEEP_BETWEEN`, `LOG_LEVEL`, `TIMEOUT_SECONDS`, `USER_AGENT` (`main-code.py:68-74`): Runtime knobs controlling lookback window, archival horizon, throttling, and HTTP settings.

### HTTP Session
- Uses `requests.Session` with a `Retry` adapter so transient 4xx/5xx responses are retried automatically.
- Adds realistic headers to avoid basic anti-bot protections.

### Scraping Helpers
- `parse_hidden_fields`: Captures ASP.NET state fields needed for paging.
- `extract_rows`: Locates the results table, tolerating multiple class names, and reads cell values.
- `find_next_postback`: Walks pagination links to determine the next `__EVENTTARGET`/`__EVENTARGUMENT` pair.
- `parse_publish_dt`: Converts the raw timestamp string to a timezone-aware `datetime` in Eastern Time.

### Data Normalization
- Every row gains a `source` label (`emma`), a deterministic `record_id`, and placeholders for optional workbook columns like `tags`.
- Publish timestamps are converted to naive datetimes (`to_excel_naive`) before writing to Excel, because the format cannot store timezone-aware objects.

### Workbook Management
- `init_workbook_if_needed` seeds the workbook with the four sheets.
- `ensure_master_table_style` enforces a single Excel Table named `tbl_opps`, including header freeze panes and striped rows.
- `auto_col_widths` heuristically sets column widths.
- `apply_status_conditional_formats` colors statuses (green for `New`, amber for `Updated`, grey for `Stale`).

### Merge Logic
- `rows_equal` compares only the business-facing columns, ensuring admin metadata (timestamps, status) does not trigger false updates.
- Newly touched rows receive `last_seen_et` and, when applicable, `first_seen_et` timestamps; all actions are logged via the `Action` dataclass.
- Stale rows are moved to `Archive` and then removed from `Master` so the active sheet stays current.

## Workbook Schema
### `MASTER_HDR`
| Column | Purpose |
| --- | --- |
| `source` | Identifier for the data origin (`emma`). |
| `record_id` | Stable key derived from the eMMA numeric ID or a hash. |
| `url` | Direct link to the eMMA opportunity. |
| `first_seen_et` | When the row first appeared in the workbook. |
| `last_seen_et` | Timestamp of the most recent run that saw this row. |
| `title` | Opportunity title scraped from the listing. |
| `agency` | Procuring agency. |
| `category` | Procurement category field. |
| `procurement_method` | Method/vehicle listed by eMMA. |
| `publish_dt_et` | Publish datetime in Eastern Time (naive for Excel). |
| `due_dt_et` | Placeholder for future capture of due dates. |
| `status` | Change classification (`New`, `Updated`, `Unchanged`, `Stale`). |
| `tags` | Optional manual tags populated later. |
| `score_bd_fit` | Optional scoring or notes. |

### Status Meanings
| Status | Trigger |
| --- | --- |
| `New` | Record not previously seen in `Master`. |
| `Updated` | Record exists but one or more business fields changed. |
| `Unchanged` | Record exists and business fields match prior run; only `last_seen_et` is refreshed. |
| `Stale` | Record not seen in the current run and `last_seen_et` is older than `STALE_AFTER_D` days; moved to `Archive`. |

## Configuration Reference
| Environment Variable | Default | Notes |
| --- | --- | --- |
| `EMMA_XLSX` | `C:\Users\hkhoshhal001\Guidehouse\...\opportunities.xlsx` | Override with a path you can write to locally. |
| `DAYS_AGO` | `0` | Scrape listings published `n` days ago (use `1` for yesterday). |
| `STALE_AFTER_D` | `7` | Rows untouched for more than this many days are archived. |
| `MAX_PAGES` | `50` | Cap on pagination depth to avoid large crawls. |
| `SLEEP_BETWEEN` | `1.0` | Seconds to wait between page requests. |
| `LOG_LEVEL` | `INFO` | Valid values: `DEBUG`, `INFO`, `WARNING`, etc. |
| `TIMEOUT_SECONDS` | `30` | HTTP request timeout per call. |
| `USER_AGENT` | `Mozilla/5.0 (compatible; MD-EmmaScraper/1.0)` | Set if you need a custom identifier. |

## Local Setup Guide
### Prerequisites
- Python 3.9+ (3.9 ensures `zoneinfo` is available; on 3.8 install `backports.zoneinfo`).
- Install dependencies in your environment:
  ```bash
  pip install requests beautifulsoup4 pandas openpyxl
  ```
  > `pandas` is imported but currently unused; you can omit it if you prefer.

### Configure Paths
- Set `EMMA_XLSX` to a writable location, for example:
  ```bash
  export EMMA_XLSX="/path/to/opportunities.xlsx"
  ```
- Optionally adjust `DAYS_AGO`, `STALE_AFTER_D`, and other variables before running.

### Run the Script
```bash
python main-code.py
```
- The script prints `Updated workbook: <path> (target day: <n> days ago)` upon completion.
- Check the console logs for retry attempts or warnings about rebuilding the workbook.

## Operational Notes
- **Timezone Handling:** All timestamps are normalized to Eastern Time using the standard `zoneinfo` database when available, falling back to naive local time otherwise.
- **Rate Limiting:** `SLEEP_BETWEEN` introduces a delay between page fetches to avoid hammering the eMMA site. Increase the delay if you encounter throttling.
- **Resilience:** The combination of retries and page fingerprinting prevents infinite paging loops and mitigates transient HTTP errors.
- **Data Integrity:** The workbook is regenerated if it becomes corrupt (`BadZipFile` or `InvalidFileException`), ensuring the process can self-heal.

## Extending the Script
- Capture additional fields such as due dates or solicitation numbers by enhancing `extract_rows` and updating `MASTER_HDR` accordingly.
- Implement logic to auto-populate `tags` or `score_bd_fit` from rules stored in the `Refs` sheet.
- Replace the ad-hoc main block with a CLI argument parser (e.g., `argparse`) to make scheduling and parameterization easier.
- Package the script as a module with tests so CI can verify scrape parsing logic on stored HTML fixtures.

## Troubleshooting
| Symptom | Likely Cause | Suggested Remedy |
| --- | --- | --- |
| `requests.exceptions.HTTPError` | eMMA returned a non-200 status that exhausted retries. | Increase `SLEEP_BETWEEN`, verify network connectivity, or inspect logs for block messages. |
| Workbook path error | `WORKBOOK_PATH` points to a non-existent or unwritable directory. | Change `EMMA_XLSX` to a valid path. |
| No rows scraped | No listings match the chosen `DAYS_AGO` date or page structure changed. | Reduce `DAYS_AGO`, inspect `emma_scrape` logs, or capture HTML for analysis. |
| Excel shows tz-aware warning | Timestamps were not converted to naive datetimes. | Ensure `to_excel_naive` is applied before writing new datetime fields. |

## Quick Reference
- **Primary entry point:** `__main__` block at the end of `main-code.py`.
- **Core functions:** `emma_scrape`, `merge_into_excel`, `ensure_master_table_style`.
- **Outputs:** Updated Excel workbook with synchronized `Master`, `Archive`, `Log`, and `Refs` worksheets.
