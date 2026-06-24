import streamlit as st
import pdfplumber
import re
import io
from openpyxl import load_workbook
from datetime import date

# --- CONFIGURATION ---
TEMPLATE_FILE = "NOKIA-FN_06232026-001_PLAID-SITE NAME_MF2_SUBCON.xlsx"

st.set_page_config(page_title="MRF Automation Tool")
st.title("🚀 Automated MRF Generator")

# 1. File Uploads
uploaded_tssr = st.file_uploader("Upload TSSR PDF", type=["pdf"])
template_file = st.file_uploader("Upload MRF Template (Excel)", type=["xlsx"])

def extract_data_from_pdf(pdf_file):
    """Extracts data from TSSR using regex. Adjust patterns based on your PDF structure."""
    with pdfplumber.open(pdf_file) as pdf:
        text = "".join([page.extract_text() for page in pdf.pages])
    
    # Example regex patterns: Update these based on the actual keywords in your TSSR
    plaid_match = re.search(r"Site ID:\s*([A-Z0-9-]+)", text)
    cable_match = re.search(r"Fiber Cable:\s*(\d+)", text)
    card_match = re.search(r"Cards:\s*(\d+)", text)
    sfp_match = re.search(r"SFP:\s*(\d+)", text)
    
    return {
        "plaid": plaid_match.group(1) if plaid_match else "N/A",
        "cable": int(cable_match.group(1)) if cable_match else 0,
        "cards": int(card_match.group(1)) if card_match else 0,
        "sfp": int(sfp_match.group(1)) if sfp_match else 0
    }

if uploaded_tssr and template_file:
    data = extract_data_from_pdf(uploaded_tssr)
    
    st.subheader("Extracted Details")
    plaid = st.text_input("PLAID", value=data['plaid'])
    site_name = st.text_input("Site Name")
    cable_len = st.number_input("Cable Length", value=data['cable'])
    card_count = st.number_input("Number of Cards", value=data['cards'])
    sfp_count = st.number_input("Number of SFPs", value=data['sfp'])

    if st.button("Generate & Download MRF"):
        # Load the template
        wb = load_workbook(template_file)
        ws = wb.active
        
        # Mapping: Adjust these cell coordinates to your specific template
        ws['D8'] = f"{plaid} - {site_name}"
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        ws['E12'] = cable_len
        ws['E13'] = card_count
        ws['E14'] = sfp_count
        
        # Save to buffer
        output = io.BytesIO()
        wb.save(output)
        
        st.download_button(
            label="📥 Download MRF File",
            data=output.getvalue(),
            file_name=f"NOKIA-FN_06232026-001_{plaid}-{site_name}_MF2_SUBCON.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )