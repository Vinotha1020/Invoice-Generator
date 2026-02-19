import streamlit as st
import pandas as pd
from invoice_core import generate_all_invoices

st.title("Frontier Invoice Generator")

jira_file = st.file_uploader("Upload Jira Dump", type=["xlsx"])
rate_file = st.file_uploader("Upload Rate Master", type=["xlsx"])

# Store outputs
if "generated" not in st.session_state:
    st.session_state.generated = False
    st.session_state.data_file = None
    st.session_state.digital_file = None
    st.session_state.frapi_file = None


selected_month = None

# Extract months from Jira
if jira_file:

    jira_df = pd.read_excel(jira_file, sheet_name="Export")

    if "WORK_DATE" in jira_df.columns:

        jira_df['WORK_DATE'] = pd.to_datetime(jira_df['WORK_DATE'], errors='coerce')
        jira_df = jira_df.dropna(subset=['WORK_DATE'])

        jira_df['Month'] = jira_df['WORK_DATE'].dt.strftime('%Y-%m')

        month_list = sorted(jira_df['Month'].unique(), reverse=True)

        selected_month = st.selectbox(
            "Select Invoice Month",
            month_list
        )

    else:
        st.error("WORK_DATE column not found in Jira dump")


# Generate invoices
if jira_file and rate_file and selected_month:

    if st.button("Generate Invoices"):

        data_file, digital_file, frapi_file = generate_all_invoices(
            jira_file,
            rate_file,
            selected_month
        )

        st.session_state.data_file = data_file
        st.session_state.digital_file = digital_file
        st.session_state.frapi_file = frapi_file
        st.session_state.generated = True


# Persist download buttons
if st.session_state.generated:

    st.success("Invoices Generated ✅")

    st.download_button(
        "Download Data Invoice",
        st.session_state.data_file,
        file_name="Invoice_Data.xlsx"
    )

    st.download_button(
        "Download Digital Invoice",
        st.session_state.digital_file,
        file_name="Invoice_Digital.xlsx"
    )

    st.download_button(
        "Download FRAPI Invoice",
        st.session_state.frapi_file,
        file_name="Invoice_Frapi.xlsx"
    )
