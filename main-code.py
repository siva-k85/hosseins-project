"""
eMMA -> Excel updater




- Scrapes Maryland eMMA public listings using requests+bs4
- Merges into one Excel workbook:
  * Master: current rows (<= 7 days old), styled table
  * Log: append-only history of actions
  * Archive: pruned (stale) rows
  * Refs: optional rules you can fill manually
- No Chrome/driver, no Selenium




Env vars you can set:
- EMMA_XLSX (default: opportunities.xlsx)
- DAYS_AGO (default: 2)       # 1=yesterday, 2=day before, etc.
- STALE_AFTER_D (default: 7)  # days after which untouched rows are archived
- MAX_PAGES (default: 50)
- SLEEP_BETWEEN (default: 1.0)
- LOG_LEVEL (default: INFO)
- TIMEOUT_SECONDS (default: 30)
- USER_AGENT (default provided)
"""




import os
import re
import time
import logging
from hashlib import blake2b
from dataclasses import dataclass
from datetime import datetime, timedelta
from urllib.parse import urljoin
from zipfile import BadZipFile




import requests
from requests.adapters import HTTPAdapter, Retry
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException


# -------------------- Config (env-overridable) --------------------




WORKBOOK_PATH = os.getenv(
    "EMMA_XLSX",
    r"C:\Users\hkhoshhal001\Guidehouse\New York Mid-Atlantic (NYMA) - 05. Scanning resources\Automated Scanning\opportunities.xlsx"
)


DAYS_AGO        = int(os.getenv("DAYS_AGO", "0"))     # 1=yesterday, 2=day before, etc.
STALE_AFTER_D   = int(os.getenv("STALE_AFTER_D", "7"))
MAX_PAGES       = int(os.getenv("MAX_PAGES", "50"))
SLEEP_BETWEEN   = float(os.getenv("SLEEP_BETWEEN", "1.0"))
LOG_LEVEL       = os.getenv("LOG_LEVEL", "INFO").upper()
TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "30"))
USER_AGENT      = os.getenv("USER_AGENT", "Mozilla/5.0 (compatible; MD-EmmaScraper/1.0)")




BASE = "https://emma.maryland.gov"
BROWSE_URL = f"{BASE}/page.aspx/en/rfp/request_browse_public"

# -------------------- Timezone (ET) --------------------
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    ET_TZ = ZoneInfo("America/New_York")
except Exception:
    ET_TZ = None




def now_et():
    return datetime.now(ET_TZ) if ET_TZ else datetime.now()




def localize_et(dt: datetime) -> datetime:
    return dt.replace(tzinfo=ET_TZ) if ET_TZ else dt




def to_excel_naive(dt):
    """Excel/openpyxl cannot store tz-aware datetimes. Return naive (no tzinfo)."""
    if dt is None:
        return None
    if isinstance(dt, datetime) and dt.tzinfo is not None:
        return dt.replace(tzinfo=None)
    return dt




# -------------------- Logging --------------------
logging.basicConfig(level=getattr(logging, LOG_LEVEL, logging.INFO),
                    format="%(asctime)s [%(levelname)s] %(message)s")




# -------------------- HTTP session --------------------
def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": USER_AGENT,
                      "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"})
    retries = Retry(total=5, backoff_factor=0.6,
                    status_forcelist=[429,500,502,503,504],
                    allowed_methods=["GET","POST"], raise_on_status=False)
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.mount("http://", HTTPAdapter(max_retries=retries))
    return s

# -------------------- eMMA scraping --------------------
TS_PATTERN = re.compile(r"\b\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+(AM|PM)\b")




def parse_hidden_fields(soup: BeautifulSoup) -> dict:
    fields = {}
    for name in ["__VIEWSTATE","__EVENTVALIDATION","__VIEWSTATEGENERATOR","__EVENTTARGET","__EVENTARGUMENT"]:
        el = soup.find("input", {"name": name})
        if el and el.has_attr("value"):
            fields[name] = el["value"]
    return fields




def extract_rows(soup: BeautifulSoup) -> list[dict]:
    table = soup.select_one("table.iv-grid-view")
    if not table:
        for t in soup.find_all("table"):
            classes = " ".join(t.get("class", []))
            if "iv-grid" in classes or "iv-grid-view" in classes or "very compact" in classes:
                table = t
                break
    if not table:
        return []




    rows = []
    tbody = table.find("tbody") or table
    for tr in tbody.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 9:
            continue
        title_cell = tds[2]
        a = title_cell.find("a")
        title = a.get_text(strip=True) if a else title_cell.get_text(strip=True)
        link = urljoin(BASE, a["href"]) if (a and a.has_attr("href")) else None




        category = tds[6].get_text(strip=True) if len(tds) > 6 else ""
        method   = tds[7].get_text(strip=True) if len(tds) > 7 else ""
        agency   = tds[8].get_text(strip=True) if len(tds) > 8 else ""




        publish_text = None
        for td in tds:
            txt = td.get_text(" ", strip=True)
            m = TS_PATTERN.search(txt)
            if m:
                publish_text = m.group(0)
                break




        rows.append({
            "title": title or "",
            "url": link,
            "category": category,
            "procurement_method": method,
            "agency": agency,
            "publish_dt_raw": publish_text or "",
        })
    return rows




def find_next_postback(soup: BeautifulSoup) -> tuple[str|None, str|None]:
    for a in soup.find_all("a", href=True):
        m = re.search(r"__doPostBack\('([^']*)','([^']*)'\)", a["href"])
        if m and (a.get_text(strip=True) or "").lower() in {"next",">",">>"}:
            return m.group(1), m.group(2)
    # Fallback to highest page number
    candidates = []
    for a in soup.find_all("a", href=True):
        m = re.search(r"__doPostBack\('([^']*)','([^']*)'\)", a["href"])
        label = a.get_text(strip=True)
        if m and label.isdigit():
            candidates.append((int(label), m.group(1), m.group(2)))
    if candidates:
        _, et, ea = sorted(candidates, key=lambda x:x[0])[-1]
        return et, ea
    return None, None




def parse_publish_dt(raw: str) -> datetime|None:
    if not raw:
        return None
    try:
        dt = datetime.strptime(raw, "%m/%d/%Y %I:%M:%S %p")
        return localize_et(dt)
    except Exception:
        return None




def emma_scrape(DAYS_AGO: int, max_pages:int=50, sleep_s:float=1.0) -> list[dict]:
    """Return a list of normalized records for the target ET date."""
    ses = make_session()
    r = ses.get(BROWSE_URL, timeout=TIMEOUT_SECONDS)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")




    all_rows = []
    pages = 0
    seen = set()
    def page_fp(s):  # guard against repeats
        t = (s.select_one("table.iv-grid-view") or s.find("table"))
        text = (t.get_text(" ", strip=True) if t else s.get_text(" ", strip=True))[:20000]
        return blake2b(text.encode(), digest_size=16).hexdigest()




    while True:
        fp = page_fp(soup)
        if fp in seen:
            break
        seen.add(fp)




        rows = extract_rows(soup)
        all_rows.extend(rows)
        pages += 1
        if pages >= max_pages:
            break




        et, ea = find_next_postback(soup)
        if not et:
            break




        fields = parse_hidden_fields(soup)
        fields["__EVENTTARGET"]  = et
        fields["__EVENTARGUMENT"] = ea
        time.sleep(sleep_s)
        r = ses.post(BROWSE_URL, data=fields, timeout=TIMEOUT_SECONDS)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")




    # Normalize and filter by target ET date
    for row in all_rows:
        row["publish_dt_et"] = parse_publish_dt(row.pop("publish_dt_raw", ""))
        row["source"] = "emma"




        # derive record_id from URL numeric id if present
        rec_id = None
        if row["url"]:
            m = re.search(r"/extranet/(\d+)", row["url"])
            if m:
                rec_id = m.group(1)
        if not rec_id:
            # fallback: hash of URL+title
            h = blake2b((row.get("url","")+row.get("title","")).encode(), digest_size=8).hexdigest()
            rec_id = f"emma_{h}"
        row["record_id"] = rec_id




        row["due_dt_et"] = None     # not scraped yet in this pass
        row["tags"] = ""            # optional (from Refs rules later)
        row["score_bd_fit"] = ""    # optional (from Refs rules later)




    target_date = (now_et().date() - timedelta(days=DAYS_AGO))
    staging = [r for r in all_rows if r.get("publish_dt_et") and r["publish_dt_et"].date() == target_date]
    return staging







# -------------------- Excel helpers --------------------
MASTER_HDR = [
    "source","record_id","url","first_seen_et","last_seen_et",
    "title","agency","category","procurement_method","publish_dt_et","due_dt_et",
    "status","tags","score_bd_fit"
]


def create_workbook_backup(original_path: str, max_backups: int = 5) -> str:
    """
    Create a timestamped backup of the workbook file.

    Args:
        original_path: Path to the original workbook file
        max_backups: Maximum number of backups to keep (default: 5)

    Returns:
        Path to the created backup file
    """
    if not os.path.exists(original_path):
        logging.debug("Original workbook '%s' doesn't exist, skipping backup", original_path)
        return ""

    try:
        # Create backup directory
        backup_dir = os.path.join(os.path.dirname(original_path), "backups")
        os.makedirs(backup_dir, exist_ok=True)

        # Generate timestamped backup filename
        base_name = os.path.splitext(os.path.basename(original_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{base_name}_backup_{timestamp}.xlsx"
        backup_path = os.path.join(backup_dir, backup_filename)

        # Copy the file
        import shutil
        shutil.copy2(original_path, backup_path)

        logging.info("Created workbook backup: %s", backup_path)

        # Cleanup old backups
        cleanup_old_backups(backup_dir, base_name, max_backups)

        return backup_path

    except (OSError, IOError) as exc:
        logging.warning("Failed to create backup of '%s': %s", original_path, exc)
        return ""


def cleanup_old_backups(backup_dir: str, base_name: str, max_backups: int) -> None:
    """
    Remove old backup files, keeping only the most recent ones.

    Args:
        backup_dir: Directory containing backup files
        base_name: Base name of the workbook (without extension)
        max_backups: Maximum number of backups to keep
    """
    try:
        if not os.path.exists(backup_dir):
            return

        # Find all backup files for this workbook
        backup_files = []

        for filename in os.listdir(backup_dir):
            if filename.startswith(f"{base_name}_backup_") and filename.endswith(".xlsx"):
                filepath = os.path.join(backup_dir, filename)
                if os.path.isfile(filepath):
                    backup_files.append((filepath, os.path.getmtime(filepath)))

        # Sort by modification time (newest first)
        backup_files.sort(key=lambda x: x[1], reverse=True)

        # Remove old backups beyond max_backups limit
        files_to_remove = backup_files[max_backups:]
        removed_count = 0

        for filepath, _ in files_to_remove:
            try:
                os.remove(filepath)
                removed_count += 1
                logging.debug("Removed old backup: %s", filepath)
            except OSError as exc:
                logging.warning("Failed to remove old backup '%s': %s", filepath, exc)

        if removed_count > 0:
            logging.info("Cleaned up %d old backup files in %s", removed_count, backup_dir)

    except OSError as exc:
        logging.warning("Failed to cleanup old backups in '%s': %s", backup_dir, exc)




def init_workbook_if_needed(path:str):
    if os.path.exists(path):
        return
    wb = Workbook()
    # create sheets in order
    ws_master = wb.active
    ws_master.title = "Master"
    ws_master.append(MASTER_HDR)




    ws_log = wb.create_sheet("Log")
    ws_log.append(["run_ts_et","action"] + MASTER_HDR)




    wb.create_sheet("Refs")
    wb.create_sheet("Archive")




    create_workbook_backup(path)
    wb.save(path)




def load_wb(path:str):
    # Create if missing
    init_workbook_if_needed(path)
    # Try opening; if corrupt, rebuild
    try:
        return load_workbook(path)
    except (BadZipFile, InvalidFileException):
        logging.warning("File at %s was not a valid Excel. Reinitializing.", path)
        try:
            os.remove(path)
        except Exception:
            pass
        init_workbook_if_needed(path)
        return load_workbook(path)




def ws_to_index(ws, key_col_name="record_id") -> dict:
    # build index of existing Master rows: record_id -> row number
    header = [c.value for c in ws[1]]
    col_idx = {name: header.index(name)+1 for name in header}
    idx = {}
    for row in range(2, ws.max_row+1):
        rid = ws.cell(row=row, column=col_idx[key_col_name]).value
        if rid:
            idx[rid] = row
    return header, col_idx, idx




def ensure_master_table_style(ws):
    """
    Ensure there is a single styled Excel Table named tbl_opps that
    spans all data in ws. If a table already exists, resize it.
    """
    last_row = ws.max_row
    last_col = ws.max_column
    if last_row < 1 or last_col < 1:
        return




    ref = f"A1:{get_column_letter(last_col)}{last_row}"




    # openpyxl exposes tables via ws.tables (dict-like). If present, resize; if not, create.
    existing_tables = getattr(ws, "tables", {})
    tbl = existing_tables.get("tbl_opps") if existing_tables else None




    if tbl:
        tbl.ref = ref
    else:
        tbl = Table(displayName="tbl_opps", ref=ref)
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        tbl.tableStyleInfo = style
        ws.add_table(tbl)




    # Freeze header row for readability
    ws.freeze_panes = "A2"




def auto_col_widths(ws, max_width=60):
    # auto size columns based on content (simple heuristic)
    for col in range(1, ws.max_column+1):
        values = [str(ws.cell(row=r, column=col).value or "") for r in range(1, ws.max_row+1)]
        width = min(max((len(v) for v in values), default=10) + 2, max_width)
        ws.column_dimensions[get_column_letter(col)].width = width




def apply_status_conditional_formats(ws):
    # Simple conditional fill by status (assumes status is in column 'L' per MASTER_HDR)
    header = [c.value for c in ws[1]]
    if "status" not in header: return
    c = header.index("status") + 1
    col_letter = get_column_letter(c)
    nrows = ws.max_row
    if nrows < 2: return




    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    amber = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    grey  = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")




    # New
    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{nrows}",
        FormulaRule(formula=[f'${col_letter}2="New"'], stopIfTrue=False, fill=green))
    # Updated
    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{nrows}",
        FormulaRule(formula=[f'${col_letter}2="Updated"'], stopIfTrue=False, fill=amber))
    # Stale
    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{nrows}",
        FormulaRule(formula=[f'${col_letter}2="Stale"'], stopIfTrue=False, fill=grey))




def row_dict(ws, row_num:int) -> dict:
    header = [c.value for c in ws[1]]
    return {name: ws.cell(row=row_num, column=i+1).value for i, name in enumerate(header)}




def rows_equal(existing:dict, incoming:dict) -> bool:
    # compare user-visible business fields only
    keys = ["title","agency","category","procurement_method","publish_dt_et","due_dt_et","url"]
    for k in keys:
        if existing.get(k) != incoming.get(k):
            return False
    return True




# -------------------- Merge pipeline --------------------
@dataclass
class Action:
    action: str
    row: dict




def merge_into_excel(staging: list[dict]):
    ts_run = now_et()
    ts_run_xl = to_excel_naive(ts_run)




    wb = load_wb(WORKBOOK_PATH)
    ws_master = wb["Master"]
    ws_log = wb["Log"]
    ws_archive = wb["Archive"]




    # Ensure headers on Master exist
    if ws_master.max_row == 1 and [c.value for c in ws_master[1]] != MASTER_HDR:
        ws_master.delete_rows(1, ws_master.max_row)
        ws_master.append(MASTER_HDR)




    header, col_idx, index = ws_to_index(ws_master)




    actions: list[Action] = []
    touched_ids = set()




    # Convert staging rows to Master schema
    for r in staging:
        row = {
            "source": "emma",
            "record_id": r["record_id"],
            "url": r["url"],
            "first_seen_et": None,  # set on insert
            "last_seen_et": ts_run_xl,
            "title": r["title"],
            "agency": r["agency"],
            "category": r["category"],
            "procurement_method": r["procurement_method"],
            "publish_dt_et": r["publish_dt_et"],
            "due_dt_et": r.get("due_dt_et"),
            "status": None,         # set below
            "tags": r.get("tags",""),
            "score_bd_fit": r.get("score_bd_fit",""),
        }




        # Ensure datetime fields are Excel-safe (naive)
        row["publish_dt_et"] = to_excel_naive(row["publish_dt_et"])
        row["due_dt_et"]     = to_excel_naive(row["due_dt_et"])




        rid = row["record_id"]
        touched_ids.add(rid)




        if rid not in index:
            # NEW
            row["first_seen_et"] = ts_run_xl
            row["status"] = "New"
            ws_master.append([row[k] for k in MASTER_HDR])
            actions.append(Action("New", row))
            # update index to include this newly added row
            index[rid] = ws_master.max_row
        else:
            # EXISTING
            existing = row_dict(ws_master, index[rid])
            row["first_seen_et"] = existing.get("first_seen_et") or ts_run_xl
            if rows_equal(existing, row):
                # Unchanged
                ws_master.cell(row=index[rid], column=col_idx["last_seen_et"]).value = ts_run_xl
                ws_master.cell(row=index[rid], column=col_idx["status"]).value = "Unchanged"
                actions.append(Action("Unchanged", row))
            else:
                # Updated fields
                for k in ["url","title","agency","category","procurement_method",
                          "publish_dt_et","due_dt_et","tags","score_bd_fit"]:
                    val = row[k]
                    if k in ("publish_dt_et","due_dt_et"):
                        val = to_excel_naive(val)
                    ws_master.cell(row=index[rid], column=col_idx[k]).value = val
                ws_master.cell(row=index[rid], column=col_idx["last_seen_et"]).value = ts_run_xl
                ws_master.cell(row=index[rid], column=col_idx["status"]).value = "Updated"
                actions.append(Action("Updated", row))




    # Prune stale: any record not touched this run AND last_seen_et older than STALE_AFTER_D days
    header, col_idx, index = ws_to_index(ws_master)
    stale_cutoff = to_excel_naive(ts_run - timedelta(days=STALE_AFTER_D))
    to_archive = []
    for rid, rownum in list(index.items()):
        if rid in touched_ids:
            continue
        last_seen = ws_master.cell(row=rownum, column=col_idx["last_seen_et"]).value
        if isinstance(last_seen, datetime) and last_seen < stale_cutoff:
            # mark stale and move to Archive
            ws_master.cell(row=rownum, column=col_idx["status"]).value = "Stale"
            arc_values = [ws_master.cell(row=rownum, column=i+1).value for i in range(len(MASTER_HDR))]
            ws_archive.append(arc_values)
            to_archive.append(rownum)
            actions.append(Action("Stale", row_dict(ws_master, rownum)))




    # Delete archived rows from Master (bottom-up)
    for rownum in sorted(to_archive, reverse=True):
        ws_master.delete_rows(rownum, 1)





    # Refresh style
    ensure_master_table_style(ws_master)
    auto_col_widths(ws_master)
    apply_status_conditional_formats(ws_master)




    # Append run actions to Log (ensure header)
    if ws_log.max_row == 1 and [c.value for c in ws_log[1]] != ["run_ts_et","action"] + MASTER_HDR:
        ws_log.delete_rows(1, ws_log.max_row)
        ws_log.append(["run_ts_et","action"] + MASTER_HDR)




    for a in actions:
        row_for_log = a.row.copy()
        for k in ("first_seen_et","last_seen_et","publish_dt_et","due_dt_et"):
            row_for_log[k] = to_excel_naive(row_for_log.get(k))
        ws_log.append([ts_run_xl, a.action] + [row_for_log.get(k) for k in MASTER_HDR])




    # Create backup before saving
    create_workbook_backup(WORKBOOK_PATH)
    wb.save(WORKBOOK_PATH)
    logging.info(
        "Workbook saved: %s | Actions -> New:%d Updated:%d Unchanged:%d Stale:%d",
        WORKBOOK_PATH,
        sum(1 for x in actions if x.action == "New"),
        sum(1 for x in actions if x.action == "Updated"),
        sum(1 for x in actions if x.action == "Unchanged"),
        sum(1 for x in actions if x.action == "Stale"),
    )




# -------------------- Entry point --------------------
if __name__ == "__main__":
    # 1) scrape staging for target ET date
    staging = emma_scrape(DAYS_AGO=DAYS_AGO, max_pages=MAX_PAGES, sleep_s=SLEEP_BETWEEN)
    # 2) merge into Excel workbook (no CSVs)
    merge_into_excel(staging)
    print(f"Updated workbook: {WORKBOOK_PATH} (target day: {DAYS_AGO} days ago)")
