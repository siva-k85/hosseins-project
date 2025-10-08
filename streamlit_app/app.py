import os
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Dict, Tuple

import pandas as pd
import streamlit as st


DEFAULT_WORKBOOK = Path(os.getenv("EMMA_XLSX", Path.home() / "Documents" / "emma_opportunities.xlsx"))

st.set_page_config(
    page_title="eMMA Opportunities Portal",
    page_icon="📊",
    layout="wide",
)

CUSTOM_CSS = """
<style>
    .main {
        background: linear-gradient(135deg, #f9fbff 0%, #eef3ff 100%);
    }
    .stMetric label, .metric-container span {
        font-weight: 700 !important;
        color: #2d4a8a !important;
    }
    div[data-testid="stMetricValue"] {
        color: #1f2c56 !important;
    }
    .download-card {
        background: #ffffffcc;
        border-radius: 18px;
        padding: 1.5rem;
        box-shadow: 0 12px 32px rgba(31,76,135,0.12);
        border: 1px solid rgba(85,112,172,0.12);
    }
    .section-title {
        font-size: 1.45rem;
        font-weight: 700;
        color: #1f2c56;
        margin-top: 1.5rem;
        margin-bottom: 0.5rem;
    }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def _resolve_workbook_path(selection: str) -> Path:
    candidate = Path(selection).expanduser()
    if candidate.is_file():
        return candidate
    return DEFAULT_WORKBOOK


@st.cache_data(show_spinner="Loading workbook…")
def load_workbook(path: str) -> Tuple[Dict[str, pd.DataFrame], datetime]:
    resolved = _resolve_workbook_path(path)
    if not resolved.exists():
        raise FileNotFoundError(f"Workbook not found at {resolved}")
    sheet_map = pd.read_excel(resolved, sheet_name=None)
    modified = datetime.fromtimestamp(resolved.stat().st_mtime)
    return sheet_map, modified


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.read()


st.sidebar.title("📁 Workbook Settings")
workbook_input = st.sidebar.text_input(
    "Workbook path", value=str(DEFAULT_WORKBOOK)
)
try:
    sheets, modified_ts = load_workbook(workbook_input)
except FileNotFoundError as exc:
    st.error(str(exc))
    st.stop()

master_df = sheets.get("Master", pd.DataFrame())
log_df = sheets.get("Log", pd.DataFrame())

st.title("Maryland eMMA Opportunities Dashboard")
st.caption("A beautiful interface for exploring and exporting procurement data scraped by the automation suite.")

col_info, col_download = st.columns([2, 1])
with col_info:
    st.subheader("Workbook Overview")
    st.write(
        "Last updated: **{}**".format(modified_ts.strftime("%B %d, %Y – %I:%M %p"))
    )
with col_download:
    with open(_resolve_workbook_path(workbook_input), "rb") as f:
        st.download_button(
            "⬇️ Download full workbook",
            data=f.read(),
            file_name=_resolve_workbook_path(workbook_input).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
if not master_df.empty:
    metric_col1.metric("Active Opportunities", len(master_df))
    metric_col2.metric("New This Run", int((master_df["status"] == "New").sum()) if "status" in master_df else 0)
    metric_col3.metric(
        "Updated", int((master_df["status"] == "Updated").sum()) if "status" in master_df else 0
    )
    upcoming_mask = pd.Series(False, index=master_df.index)
    if "due_dt_et" in master_df:
        upcoming_mask = pd.to_datetime(master_df["due_dt_et"], errors="coerce") <= datetime.now() + timedelta(days=7)
    metric_col4.metric("Due Within 7 Days", int(upcoming_mask.sum()))
else:
    metric_col1.metric("Active Opportunities", 0)
    metric_col2.metric("New This Run", 0)
    metric_col3.metric("Updated", 0)
    metric_col4.metric("Due Within 7 Days", 0)

st.markdown("<div class='section-title'>🔎 Explore Opportunities</div>", unsafe_allow_html=True)

explorer_col1, explorer_col2 = st.columns([2, 1])
with explorer_col1:
    sheet_choice = st.selectbox("Select sheet", options=list(sheets.keys()), index=0 if "Master" in sheets else 0)
with explorer_col2:
    search_term = st.text_input("Search title or agency", "")

active_df = sheets[sheet_choice].copy()

if search_term:
    lowered = search_term.lower()
    active_df = active_df[active_df.apply(lambda row: row.astype(str).str.lower().str.contains(lowered, na=False).any(), axis=1)]

if sheet_choice == "Master" and "due_dt_et" in active_df:
    date_min = pd.to_datetime(active_df["due_dt_et"], errors="coerce").min()
    date_max = pd.to_datetime(active_df["due_dt_et"], errors="coerce").max()
    if pd.notnull(date_min) and pd.notnull(date_max):
        start, end = st.slider(
            "Due date window",
            value=(date_min.to_pydatetime(), date_max.to_pydatetime()),
            min_value=date_min.to_pydatetime(),
            max_value=date_max.to_pydatetime(),
        )
        mask = pd.to_datetime(active_df["due_dt_et"], errors="coerce").between(start, end)
        active_df = active_df[mask]

st.dataframe(active_df, use_container_width=True, height=480)

st.markdown("<div class='section-title'>📥 Download Options</div>", unsafe_allow_html=True)
download_col1, download_col2, download_col3 = st.columns(3)
with download_col1:
    st.download_button(
        "Download current view",
        data=to_excel_bytes(active_df, sheet_choice),
        file_name=f"emma_{sheet_choice.lower()}_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with download_col2:
    if sheet_choice != "Master" and "Master" in sheets:
        st.download_button(
            "Download Master sheet",
            data=to_excel_bytes(sheets["Master"], "Master"),
            file_name="emma_master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
with download_col3:
    if sheet_choice != "Log" and "Log" in sheets:
        st.download_button(
            "Download Log sheet",
            data=to_excel_bytes(sheets["Log"], "Log"),
            file_name="emma_log.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if not log_df.empty and {"run_ts_et", "action"}.issubset(log_df.columns):
    st.markdown("<div class='section-title'>📈 Recent Run Activity</div>", unsafe_allow_html=True)
    chart_data = (
        log_df[["run_ts_et", "action"]]
        .groupby(["run_ts_et", "action"])  # type: ignore[arg-type]
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )
    chart_data["run_ts_et"] = pd.to_datetime(chart_data["run_ts_et"], errors="coerce")
    chart_data = chart_data.sort_values("run_ts_et", ascending=True).tail(20)
    chart_data.set_index("run_ts_et", inplace=True)
    st.area_chart(chart_data)

st.markdown(
    """
    <div class='download-card'>
        <h3 style="margin-bottom:0.25rem; color:#1f2c56;">Tips</h3>
        <ul style="margin-top:0.5rem;">
            <li>Use the sidebar to point to a different workbook location.</li>
            <li>Apply search and date filters before exporting the filtered view.</li>
            <li>Switch between sheets (Master, Log, Archive, Refs) using the selector.</li>
        </ul>
    </div>
    """,
    unsafe_allow_html=True,
)
