import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import date

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

# 1. Load Database
@st.cache_data
def load_db():
    try:
        # Simplified filename to avoid FileNotFoundError
        return pd.read_csv("eul_material_list.csv").set_index('PLAID')
    except FileNotFoundError:
        st.error("Database file 'eul_material_list.csv' not found!")
        return None

material_db = load_db()

if material_db is not None:
    # 2. Site Selection
    selected_plaid = st.sidebar.selectbox("Select PLAID", material_db.index.tolist())
    site_name = st.sidebar.text_input("Enter Site Name")

    if st.button("Generate MRF"):
        site_row = material_db.loc[selected_plaid]
        
        # 3. Load Template
        wb = load_workbook("template_mrf.xlsx")
        ws = wb.active
        
        # 4. Fill Metadata
        ws['D8'] = f"{selected_plaid} - {site_name}"
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        
        # 5. Automated Mapping: Loop through Template rows (16 to 60)
        for row in range(16, 60):
            part_no = str(ws[f'A{row}'].value).strip()
            
            # Check if this part number exists in our CSV database
            if part_no in material_db.columns:
                quantity = site_row[part_no]
                ws[f'D{row}'] = quantity
        
        # 6. Save and Download
        output = io.BytesIO()
        wb.save(output)
        
        st.success(f"MRF generated for {selected_plaid}!")
        st.download_button("📥 Download MRF", data=output.getvalue(), 
                           file_name=f"MRF_{selected_plaid}.xlsx")