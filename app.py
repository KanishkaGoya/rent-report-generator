"""
app.py

Streamlit front-end for the Rent Report Generator application.

This app allows a user to:
    1. Upload a Vendor Master Excel file.
    2. Upload a Rent Report Excel file.
    3. Enter Company Name, Assessment Year, and Report Title.
    4. Click "Generate Report" to produce a formatted Rent Report Excel
       workbook (with a Missing_Vendors sheet for unmatched suppliers).
    5. Download the resulting file.

No data is persisted to disk - all processing happens in-memory using
BytesIO, making the app safe to run on Streamlit Community Cloud.
"""

import traceback

import streamlit as st

from report_generator import (
    RentReportError,
    process_expense_report,
)


# --------------------------------------------------------------------------
# Page Configuration
# --------------------------------------------------------------------------

st.set_page_config(
    page_title="Expense Report Generator",
    page_icon="📄",
    layout="centered",
)


# --------------------------------------------------------------------------
# Default Values
# --------------------------------------------------------------------------

DEFAULT_COMPANY_NAME = "NATIONAL ENGINEERING INDUSTRIES LIMITED"
DEFAULT_ASSESSMENT_YEAR = "2025-26"
DEFAULT_REPORT_TITLE = "DETAIL OF EXPENSES"


# --------------------------------------------------------------------------
# Session State Initialization
# --------------------------------------------------------------------------

if "generated_excel_bytes" not in st.session_state:
    st.session_state.generated_excel_bytes = None

if "matched_count" not in st.session_state:
    st.session_state.matched_count = None

if "missing_count" not in st.session_state:
    st.session_state.missing_count = None


# --------------------------------------------------------------------------
# UI - Title
# --------------------------------------------------------------------------

st.title("Expense Report Generator")
st.write(
    "Upload the Vendor Master and Expense Report Excel files, fill in the "
    "report details below, and click **Generate Report** to produce a "
    "formatted, downloadable Expense Report."
)

st.divider()

# --------------------------------------------------------------------------
# UI - File Upload Controls
# --------------------------------------------------------------------------

st.subheader("1. Upload Files")

vendor_master_file = st.file_uploader(
    "Upload Vendor Master (.xlsx)",
    type=["xlsx"],
    key="vendor_master_uploader",
    help=(
        "Excel file containing vendor master data with columns: Vendor, "
        "City, Name 1, GST Registration No., Name 2, Street, Name 3, "
        "Name 4, PostalCode, PAN."
    ),
)

expense_report_file = st.file_uploader(
    "Upload Expense Report (.xlsx)",
    type=["xlsx"],
    key="rent_report_uploader",
    help=(
        "Excel file containing expense transaction data. Only the 'Supplier' "
        "and 'Amount' columns will be used; all other columns are ignored."
    ),
)

st.divider()

# --------------------------------------------------------------------------
# UI - Text Inputs
# --------------------------------------------------------------------------

st.subheader("2. Report Details")

company_name = st.text_input(
    "Company Name",
    value=DEFAULT_COMPANY_NAME,
)

assessment_year = st.text_input(
    "Assessment Year",
    value=DEFAULT_ASSESSMENT_YEAR,
)

report_title = st.text_input(
    "Report Title",
    value=DEFAULT_REPORT_TITLE,
)

st.divider()

# --------------------------------------------------------------------------
# UI - Generate Button & Processing
# --------------------------------------------------------------------------

st.subheader("3. Generate Report")

generate_clicked = st.button("Generate Report", type="primary")

if generate_clicked:
    # Reset any previous results before attempting a new generation
    st.session_state.generated_excel_bytes = None
    st.session_state.matched_count = None
    st.session_state.missing_count = None

    # --- Input presence validation ---
    if vendor_master_file is None or expense_report_file is None:
        st.error(
            "Please upload both the Vendor Master and Expense Report Excel "
            "files before generating the report."
        )
    elif not company_name.strip():
        st.error("Company Name cannot be empty.")
    elif not assessment_year.strip():
        st.error("Assessment Year cannot be empty.")
    elif not report_title.strip():
        st.error("Report Title cannot be empty.")
    else:
        try:
            with st.spinner("Generating report, please wait..."):
                excel_buffer, matched_count, missing_count = process_rent_report(
                    vendor_master_file=vendor_master_file,
                    rent_report_file=rent_report_file,
                    company_name=company_name,
                    assessment_year=assessment_year,
                    report_title=report_title,
                )

            # Persist results in session state so the download button
            # survives Streamlit's rerun-on-interaction behavior.
            st.session_state.generated_excel_bytes = excel_buffer.getvalue()
            st.session_state.matched_count = matched_count
            st.session_state.missing_count = missing_count

            st.success("Report generated successfully!")

            if missing_count > 0:
                st.warning(
                    f"{missing_count} supplier(s) from the Expense Report "
                    f"could not be matched to a Vendor Master entry. "
                    f"They have been listed in the 'Missing_Vendors' sheet "
                    f"of the generated file."
                )

        except RentReportError as known_error:
            # Friendly, user-facing errors raised deliberately by the
            # report_generator module (missing columns, empty files, etc.)
            st.error(str(known_error))

        except Exception:  # noqa: BLE001 - final safety net, never crash
            st.error(
                "An unexpected error occurred while generating the report. "
                "Please check your input files and try again."
            )
            with st.expander("Show technical details"):
                st.code(traceback.format_exc())

# --------------------------------------------------------------------------
# UI - Download Button (shown whenever a report has been generated)
# --------------------------------------------------------------------------

if st.session_state.generated_excel_bytes is not None:
    st.divider()
    st.subheader("4. Download Report")

    matched_count = st.session_state.matched_count
    missing_count = st.session_state.missing_count

    st.write(
        f"**Matched suppliers:** {matched_count}  \n"
        f"**Missing suppliers:** {missing_count}"
    )
        download_filename = (
           report_title.strip()
           .replace(" ", "_")
           .replace("/", "-")
           + ".xlsx"
    )
    st.download_button(
        label="Download Final Report.xlsx",
        data=st.session_state.generated_excel_bytes,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
