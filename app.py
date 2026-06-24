import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import date

# 1. Load Data
@st.cache_data
def load_db():
    # Load the Material List
    df = pd.read_csv("MINDANAO _Site_Activity_Monitoring_OLT PROJECT.xlsx - EUL material list.csv")
    return df.set_index('PLAID')

# 2. Page Config
st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

material_db = load_db()

# 3. Sidebar Selection
st.sidebar.header("Site Selection")
selected_plaid = st.sidebar.selectbox("Select PLAID", material_db.index.tolist())
site_name = st.sidebar.text_input("Enter Site Name for Output")

# 4. Processing
if st.button("Generate MRF"):
    # Get quantities for selected site
    site_row = material_db.loc[selected_plaid]
    
    # Load Template
    template_file = "NOKIA-FN_06232026-001_PLAID-SITE NAME_MF2_SUBCON.xlsx"
    wb = load_workbook(template_file)
    ws = wb.active
    
    # Update Header Metadata
    ws['D8'] = f"{selected_plaid} - {site_name}"
    ws['E4'] = date.today().strftime("%Y-%m-%d")
    
    # 5. Mapping Logic: Scan rows for Part Numbers and update from CSV
    # Template assumes Part Number in Column A, Quantity in Column D (index 4)
    for row in range(16, 60):  # Adjust range to cover your item rows
        part_no = ws[f'A{row}'].value
        
        # If the cell contains a part number that exists in our CSV columns
        if part_no in site_row.index:
            ws[f'D{row}'] = site_row[part_no]
            
    # Save to buffer
    output = io.BytesIO()
    wb.save(output)
    
    st.success(f"MRF generated for {selected_plaid}!")
    st.download_button(
        label="📥 Download MRF File",
        data=output.getvalue(),
        file_name=f"NOKIA-FN_06232026-001_{selected_plaid}-{site_name}_MF2_SUBCON.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Visualizing the Data Flow
st.info("The application automatically matches the Part Numbers in the Excel Template with the columns in your EUL Material List CSV.")