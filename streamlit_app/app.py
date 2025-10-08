import streamlit as st
import pandas as pd
import os
import sys
from io import StringIO, BytesIO
import logging
from datetime import datetime, timedelta
import json
import plotly.express as px
import plotly.graph_objects as go

# Add project root to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Import the scraper - handle both main-code.py and main_code.py
run_scraper = None
try:
    import importlib.util
    main_code_path = os.path.join(os.path.dirname(__file__), '..', 'main-code.py')
    if os.path.exists(main_code_path):
        spec = importlib.util.spec_from_file_location("main_code", main_code_path)
        if spec and spec.loader:
            main_code = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(main_code)
            run_scraper = main_code.main
except Exception as e:
    pass

# Fallback: try importing from emma_scraper_consolidated
if run_scraper is None:
    try:
        from emma_scraper_consolidated import main as run_scraper
    except:
        try:
            # Final fallback: try emma_scraper_ultimate
            from emma_scraper_ultimate import main as run_scraper
        except:
            run_scraper = None

# --- Page Configuration ---
st.set_page_config(
    page_title="eMMA Scraper Dashboard",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- Styling ---
st.markdown("""
<style>
    .reportview-container {
        background: #f0f2f6;
    }
    .sidebar .sidebar-content {
        background: #ffffff;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border-radius: 12px;
        padding: 10px 24px;
        border: none;
        font-size: 16px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #45a049;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        transform: translateY(-2px);
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .urgent-card {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .info-card {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .success-card {
        background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    h1 {
        color: #1f2937;
        font-weight: 700;
    }
    h2, h3 {
        color: #374151;
        font-weight: 600;
    }
    .dataframe {
        font-size: 14px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding: 0 24px;
        background-color: #f3f4f6;
        border-radius: 8px 8px 0 0;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4CAF50;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# --- Helper Functions ---
@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_excel_data(workbook_path):
    """Load all sheets from Excel with caching"""
    try:
        excel_file = pd.ExcelFile(workbook_path)
        sheets = {}
        for sheet_name in excel_file.sheet_names:
            sheets[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name)
        return sheets, None
    except Exception as e:
        return None, str(e)

def create_metrics_cards(df):
    """Create beautiful metric cards"""
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("Total Opportunities", len(df))
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        urgent = len(df[df['days_until_due'] <= 7]) if 'days_until_due' in df.columns else 0
        st.markdown('<div class="urgent-card">', unsafe_allow_html=True)
        st.metric("Urgent (≤7 days)", urgent)
        st.markdown('</div>', unsafe_allow_html=True)

    with col3:
        open_opps = len(df[df['status'] == 'Open']) if 'status' in df.columns else 0
        st.markdown('<div class="success-card">', unsafe_allow_html=True)
        st.metric("Open Opportunities", open_opps)
        st.markdown('</div>', unsafe_allow_html=True)

    with col4:
        agencies = df['issuing_agency'].nunique() if 'issuing_agency' in df.columns else 0
        st.markdown('<div class="info-card">', unsafe_allow_html=True)
        st.metric("Unique Agencies", agencies)
        st.markdown('</div>', unsafe_allow_html=True)

def create_visualizations(df):
    """Create interactive charts"""
    if df.empty:
        st.warning("No data available for visualization")
        return

    col1, col2 = st.columns(2)

    with col1:
        if 'procurement_type' in df.columns:
            st.subheader("📊 Opportunities by Procurement Type")
            type_counts = df['procurement_type'].value_counts().head(10)
            fig = px.bar(
                x=type_counts.values,
                y=type_counts.index,
                orientation='h',
                labels={'x': 'Count', 'y': 'Procurement Type'},
                color=type_counts.values,
                color_continuous_scale='Viridis'
            )
            fig.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if 'issuing_agency' in df.columns:
            st.subheader("🏛️ Top Agencies by Opportunities")
            agency_counts = df['issuing_agency'].value_counts().head(10)
            fig = px.pie(
                values=agency_counts.values,
                names=agency_counts.index,
                hole=0.4
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

    # Timeline chart
    if 'response_deadline' in df.columns:
        st.subheader("📅 Opportunities Timeline")
        try:
            df_timeline = df.copy()
            df_timeline['response_deadline'] = pd.to_datetime(df_timeline['response_deadline'], errors='coerce')
            df_timeline = df_timeline.dropna(subset=['response_deadline'])

            if not df_timeline.empty:
                deadline_counts = df_timeline.groupby(df_timeline['response_deadline'].dt.date).size().reset_index()
                deadline_counts.columns = ['Date', 'Count']

                fig = px.line(
                    deadline_counts,
                    x='Date',
                    y='Count',
                    markers=True,
                    labels={'Count': 'Number of Deadlines', 'Date': 'Date'}
                )
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Could not create timeline chart: {e}")

def filter_dataframe(df):
    """Add filters to dataframe"""
    with st.expander("🔍 Filter Options", expanded=False):
        col1, col2, col3 = st.columns(3)

        filters = {}

        with col1:
            if 'status' in df.columns:
                statuses = ['All'] + list(df['status'].unique())
                selected_status = st.selectbox("Status", statuses)
                if selected_status != 'All':
                    filters['status'] = selected_status

        with col2:
            if 'procurement_type' in df.columns:
                proc_types = ['All'] + list(df['procurement_type'].unique())
                selected_type = st.selectbox("Procurement Type", proc_types)
                if selected_type != 'All':
                    filters['procurement_type'] = selected_type

        with col3:
            if 'issuing_agency' in df.columns:
                agencies = ['All'] + sorted(list(df['issuing_agency'].unique()))
                selected_agency = st.selectbox("Issuing Agency", agencies)
                if selected_agency != 'All':
                    filters['issuing_agency'] = selected_agency

        # Search box
        search_term = st.text_input("🔎 Search in Title or Description", "")

    # Apply filters
    filtered_df = df.copy()
    for column, value in filters.items():
        filtered_df = filtered_df[filtered_df[column] == value]

    if search_term:
        if 'opportunity_title' in filtered_df.columns:
            mask = filtered_df['opportunity_title'].str.contains(search_term, case=False, na=False)
            if 'additional_information' in filtered_df.columns:
                mask |= filtered_df['additional_information'].astype(str).str.contains(search_term, case=False, na=False)
            filtered_df = filtered_df[mask]

    return filtered_df

# --- Sidebar ---
with st.sidebar:
    st.title("🎮 eMMA Scraper Controls")
    st.markdown("---")

    # Mode selection
    mode = st.radio(
        "Select Mode",
        ["📊 View Existing Data", "🚀 Run New Scrape"],
        help="View existing Excel file or run a new scrape"
    )

    st.markdown("---")

    if mode == "🚀 Run New Scrape":
        st.subheader("Scraper Settings")
        days_ago = st.number_input(
            "Days Ago to Scrape",
            min_value=0,
            max_value=30,
            value=0,
            help="0 for today, 1 for yesterday, etc."
        )
        max_pages = st.slider(
            "Max Pages to Scrape",
            min_value=1,
            max_value=20,
            value=1,
            help="Maximum number of pages to scrape"
        )
        skip_details = st.checkbox(
            "Skip Detail Pages",
            value=False,
            help="Faster, but less data per opportunity."
        )
    else:
        st.subheader("Data Settings")
        workbook_path = st.text_input(
            "Excel File Path",
            value=os.getenv("EMMA_XLSX", "consolidated_opportunities.xlsx"),
            help="Path to the Excel file to view"
        )
        auto_refresh = st.checkbox("Auto-refresh data", value=False, help="Refresh data every 5 minutes")

    st.markdown("---")
    st.info("💡 **Tip:** Use View mode to explore existing data, or Run mode to scrape new opportunities.")

    # Statistics
    st.markdown("### 📈 Quick Stats")
    if os.path.exists(os.getenv("EMMA_XLSX", "consolidated_opportunities.xlsx")):
        file_size = os.path.getsize(os.getenv("EMMA_XLSX", "consolidated_opportunities.xlsx")) / 1024  # KB
        st.metric("File Size", f"{file_size:.1f} KB")
        mod_time = datetime.fromtimestamp(os.path.getmtime(os.getenv("EMMA_XLSX", "consolidated_opportunities.xlsx")))
        st.metric("Last Modified", mod_time.strftime("%Y-%m-%d %H:%M"))

# --- Main Page ---
st.title("🤖 eMMA Procurement Opportunity Scraper")
st.markdown("### Your automated tool for tracking Maryland's public procurement landscape.")

if mode == "🚀 Run New Scrape":
    # --- Scraper Mode ---

    # Check if scraper is available
    if run_scraper is None:
        st.error("❌ Scraper module not found. Please ensure main-code.py exists in the project root.")
        st.info("💡 Available files: Check that main-code.py or emma_scraper_consolidated.py is present.")
    elif st.button("▶️ Run Scraper and Generate Excel File", use_container_width=True):

        # --- Capture logs ---
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        redirected_output = sys.stdout = StringIO()
        sys.stderr = redirected_output

        try:
            progress_bar = st.progress(0)
            status_text = st.empty()

            status_text.text("Initializing scraper...")
            progress_bar.progress(10)

            with st.spinner("Scraping in progress... This may take a few minutes."):
                # Set environment variables for the scraper
                os.environ['DAYS_AGO'] = str(days_ago)
                if 'max_pages' in locals():
                    os.environ['MAX_PAGES'] = str(max_pages)

                # Prepare command-line arguments for argparse
                original_argv = sys.argv.copy()
                sys.argv = ['streamlit_app']
                sys.argv.append(f'--days-ago={days_ago}')
                if skip_details:
                    sys.argv.append('--skip-details')
                sys.argv.append('--log-level=INFO')

                # Run the scraper
                status_text.text("Fetching opportunities...")
                progress_bar.progress(30)
                run_scraper()
                progress_bar.progress(90)

                # Restore original argv
                sys.argv = original_argv

            progress_bar.progress(100)
            status_text.text("Complete!")
            st.success("✅ Scraping and Excel file generation complete!")

            # --- Display run summary ---
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            output = redirected_output.getvalue()

            with st.expander("📋 View Scraper Logs", expanded=False):
                st.code(output, language="log")

            # --- Provide download link ---
            workbook_path = os.getenv("EMMA_XLSX", "consolidated_opportunities.xlsx")
            if os.path.exists(workbook_path):
                col1, col2 = st.columns([2, 1])

                with col1:
                    st.subheader("📥 Download Your Excel File")
                    with open(workbook_path, "rb") as fp:
                        st.download_button(
                            label="📥 Download opportunities.xlsx",
                            data=fp,
                            file_name=f"opportunities_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                with col2:
                    st.metric("File Size", f"{os.path.getsize(workbook_path) / 1024:.1f} KB")

                # --- Preview Data ---
                try:
                    sheets, error = load_excel_data(workbook_path)
                    if sheets and 'Master' in sheets:
                        df = sheets['Master']
                        st.markdown("---")
                        create_metrics_cards(df)
                        st.markdown("---")
                        st.subheader("📊 Quick Preview")
                        st.dataframe(df.head(10), use_container_width=True)
                    else:
                        st.warning(f"Could not load Master sheet: {error}")
                except Exception as e:
                    st.warning(f"Could not generate a preview: {e}")

            else:
                st.error("❌ Could not find the generated Excel file.")

        except Exception as e:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            st.error(f"❌ An error occurred during the scraping process: {e}")
            with st.expander("View Error Details"):
                st.code(redirected_output.getvalue(), language="log")
                import traceback
                st.code(traceback.format_exc(), language="python")

else:
    # --- View Mode ---
    workbook_path = os.getenv("EMMA_XLSX", "consolidated_opportunities.xlsx")

    if not os.path.exists(workbook_path):
        st.warning(f"⚠️ Excel file not found at: {workbook_path}")
        st.info("💡 Run a scrape first or check the file path in the sidebar.")
    else:
        try:
            sheets, error = load_excel_data(workbook_path)

            if error:
                st.error(f"❌ Error loading Excel file: {error}")
            else:
                # Main tabs
                tab1, tab2, tab3, tab4 = st.tabs(["📊 Dashboard", "📋 Data Explorer", "🔍 Advanced Search", "📈 Analytics"])

                with tab1:
                    # Dashboard
                    if 'Master' in sheets:
                        df = sheets['Master']

                        create_metrics_cards(df)
                        st.markdown("---")

                        # Urgent opportunities
                        if 'days_until_due' in df.columns:
                            urgent_df = df[df['days_until_due'] <= 7].sort_values('days_until_due')
                            if not urgent_df.empty:
                                st.subheader("🚨 Urgent Opportunities (Due within 7 days)")
                                display_cols = ['opportunity_title', 'issuing_agency', 'response_deadline', 'days_until_due', 'status']
                                display_cols = [col for col in display_cols if col in urgent_df.columns]
                                st.dataframe(
                                    urgent_df[display_cols].head(10),
                                    use_container_width=True,
                                    hide_index=True
                                )

                        st.markdown("---")
                        create_visualizations(df)
                    else:
                        st.warning("Master sheet not found in the Excel file.")

                with tab2:
                    # Data Explorer
                    st.subheader("📋 Browse All Sheets")

                    sheet_name = st.selectbox("Select Sheet", list(sheets.keys()))
                    df = sheets[sheet_name]

                    st.info(f"📊 {len(df)} records in {sheet_name}")

                    # Apply filters
                    filtered_df = filter_dataframe(df)

                    if len(filtered_df) != len(df):
                        st.success(f"✅ Filtered to {len(filtered_df)} records")

                    # Display data
                    st.dataframe(filtered_df, use_container_width=True, height=600)

                    # Download filtered data
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)

                    st.download_button(
                        label=f"📥 Download Filtered {sheet_name}",
                        data=buffer.getvalue(),
                        file_name=f"filtered_{sheet_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                with tab3:
                    # Advanced Search
                    st.subheader("🔍 Advanced Search")

                    if 'Master' in sheets:
                        df = sheets['Master']

                        col1, col2 = st.columns([2, 1])

                        with col1:
                            search_query = st.text_input("🔎 Search Query", placeholder="Enter keywords...")

                        with col2:
                            search_fields = st.multiselect(
                                "Search In",
                                options=['opportunity_title', 'issuing_agency', 'category', 'additional_information'],
                                default=['opportunity_title', 'issuing_agency']
                            )

                        if search_query and search_fields:
                            mask = pd.Series([False] * len(df))
                            for field in search_fields:
                                if field in df.columns:
                                    mask |= df[field].astype(str).str.contains(search_query, case=False, na=False)

                            search_results = df[mask]

                            st.success(f"Found {len(search_results)} results")

                            if not search_results.empty:
                                st.dataframe(search_results, use_container_width=True, height=500)

                                # Export results
                                buffer = BytesIO()
                                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                    search_results.to_excel(writer, sheet_name='Search Results', index=False)

                                st.download_button(
                                    label="📥 Download Search Results",
                                    data=buffer.getvalue(),
                                    file_name=f"search_results_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        else:
                            st.info("👆 Enter a search query and select fields to search")
                    else:
                        st.warning("Master sheet not found")

                with tab4:
                    # Analytics
                    st.subheader("📈 Detailed Analytics")

                    if 'Master' in sheets:
                        df = sheets['Master']

                        col1, col2 = st.columns(2)

                        with col1:
                            if 'category' in df.columns:
                                st.subheader("Categories Distribution")
                                cat_counts = df['category'].value_counts()
                                fig = px.bar(cat_counts, orientation='h')
                                st.plotly_chart(fig, use_container_width=True)

                        with col2:
                            if 'data_quality_score' in df.columns:
                                st.subheader("Data Quality Distribution")
                                fig = px.histogram(df, x='data_quality_score', nbins=20)
                                st.plotly_chart(fig, use_container_width=True)

                        # Status breakdown
                        if 'status' in df.columns:
                            st.subheader("Status Breakdown")
                            status_df = df['status'].value_counts().reset_index()
                            status_df.columns = ['Status', 'Count']
                            fig = px.bar(status_df, x='Status', y='Count', color='Status')
                            st.plotly_chart(fig, use_container_width=True)

                        # Log sheet analysis
                        if 'Log' in sheets:
                            st.subheader("📋 Recent Activity Log")
                            log_df = sheets['Log']
                            st.dataframe(log_df.tail(20), use_container_width=True)
                    else:
                        st.warning("Master sheet not found")

        except Exception as e:
            st.error(f"❌ Error: {e}")
            import traceback
            with st.expander("View Error Details"):
                st.code(traceback.format_exc())

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #6b7280;'>
    <p>Developed with ❤️ for streamlining procurement tracking |
    <a href='https://github.com' style='color: #4CAF50;'>Documentation</a> |
    <a href='https://github.com' style='color: #4CAF50;'>Report Issues</a>
    </p>
</div>
""", unsafe_allow_html=True)