import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import date
import io

st.title("MRF Generator")

# 1. File Upload
uploaded_file = st.file_uploader("Upload your MRF Template", type=["xlsx"])

if uploaded_file:
    plaid = st.text_input("Enter PLAID")
    site_name = st.text_input("Enter Site Name")
    
    # Inputs for variables
    cable_len = st.number_input("Cable Length", value=50)
    card_count = st.number_input("Number of Cards", value=2)
    sfp_count = st.number_input("Number of SFPs", value=4)

    if st.button("Generate MRF"):
        # Load the template
        wb = load_workbook(uploaded_file)
        ws = wb.active
        
        # Update fields
        ws['D8'] = f"{plaid} - {site_name}"
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        ws['E12'] = cable_len
        ws['E13'] = card_count
        ws['E14'] = sfp_count
        
        # Save to a virtual buffer
        output = io.BytesIO()
        wb.save(output)
        
        # Download button
        st.download_button(
            label="Download Updated MRF",
            data=output.getvalue(),
            file_name=f"NOKIA-FN_06232026-001_{plaid}-{site_name}_MF2_SUBCON.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )