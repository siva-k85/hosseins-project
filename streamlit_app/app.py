import streamlit as st
import pandas as pd
import os
import sys
from io import StringIO
import logging

# Add project root to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from main_code import main as run_scraper

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
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
</style>
""", unsafe_allow_html=True)

# --- Sidebar ---
with st.sidebar:
    st.title("eMMA Scraper Controls")
    st.markdown("---")
    days_ago = st.number_input(
        "Days Ago to Scrape", 
        min_value=0, 
        max_value=30, 
        value=0, 
        help="0 for today, 1 for yesterday, etc."
    )
    skip_details = st.checkbox(
        "Skip Detail Pages", 
        value=False, 
        help="Faster, but less data per opportunity."
    )
    st.markdown("---")
    st.info("This app scrapes procurement opportunities from the eMMA Maryland website and provides them as a downloadable Excel file.")

# --- Main Page ---
st.title("eMMA Procurement Opportunity Scraper")
st.markdown("### Your automated tool for tracking Maryland's public procurement landscape.")

if st.button("▶️ Run Scraper and Generate Excel File"):
    
    # --- Capture logs ---
    log_stream = StringIO()
    
    # Temporarily redirect stdout to capture the print output of the scraper
    old_stdout = sys.stdout
    redirected_output = sys.stdout = StringIO()

    try:
        with st.spinner("Scraping in progress... This may take a few minutes."):
            # Prepare arguments for the scraper function
            class Args:
                def __init__(self):
                    self.days_ago = days_ago
                    self.skip_details = skip_details
                    self.log_level = "INFO"

            # Run the scraper
            run_scraper(Args())

        st.success("✅ Scraping and Excel file generation complete!")

        # --- Display run summary ---
        sys.stdout = old_stdout # Restore stdout
        output = redirected_output.getvalue()
        st.subheader("Run Summary")
        st.code(output, language="log")

        # --- Provide download link ---
        workbook_path = os.getenv("EMMA_XLSX", "opportunities.xlsx")
        if os.path.exists(workbook_path):
            st.subheader("Download Your Excel File")
            with open(workbook_path, "rb") as fp:
                st.download_button(
                    label="📥 Download opportunities.xlsx",
                    data=fp,
                    file_name="opportunities.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            
            # --- Preview Data ---
            try:
                df = pd.read_excel(workbook_path, sheet_name="Master")
                st.subheader("Master Sheet Preview")
                st.dataframe(df.head())
            except Exception as e:
                st.warning(f"Could not generate a preview of the Excel file. Error: {e}")

        else:
            st.error("❌ Could not find the generated Excel file.")

    except Exception as e:
        sys.stdout = old_stdout # Restore stdout
        st.error(f"An error occurred during the scraping process: {e}")
        st.code(redirected_output.getvalue(), language="log")

st.markdown("---")
st.markdown("Developed with ❤️ for streamlining procurement tracking.")