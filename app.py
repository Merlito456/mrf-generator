import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import date

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

@st.cache_data
def load_data():
    # 1. Load Material List: Header in row 2 (index 1). Index=0 is Site ID.
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1', engine='python', on_bad_lines='skip', index_col=0)
    
    # 2. Load Masterlist: First column is Site ID/PLAID.
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
    master_df.rename(columns={master_df.columns[0]: 'PLAID'}, inplace=True)
    master_df.set_index('PLAID', inplace=True)
    
    # Merge Site Name
    df = mat_df.join(master_df[['SITE']], how='left')
    return df

try:
    material_db = load_data()
    
    # Sidebar Selection
    plaid_list = material_db.index.dropna().unique().tolist()
    selected_plaid = st.sidebar.selectbox("Select Site ID", plaid_list)
    site_info = material_db.loc[selected_plaid]
    site_name = site_info['SITE'] if 'SITE' in site_info else "Unknown"

    if st.button("Generate MRF"):
        wb = load_workbook("template_mrf.xlsx")
        ws = wb.active
        
        # --- PRECISE MAPPING ---
        # Destination Code (Cell D8)
        ws['D8'] = f"{selected_plaid} - {site_name}"
        # Site ID (Targeting D10 based on your snippet)
        ws['D10'] = selected_plaid
        # Date
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        
        # Mapping Quantities: Scan rows 16 to 60, Column A to D
        for row in range(16, 60):
            part_no = str(ws[f'A{row}'].value).strip() if ws[f'A{row}'].value else ""
            if part_no in material_db.columns:
                val = site_info[part_no]
                if pd.notna(val) and val != '':
                    ws[f'D{row}'] = val
        
        output = io.BytesIO()
        wb.save(output)
        
        # Naming Format
        filename = f"NOKIA-FN_06232026-001_{selected_plaid}-{site_name}_MF2_SUBCON.xlsx"
        
        st.success("MRF Generated!")
        st.download_button("📥 Download File", data=output.getvalue(), file_name=filename)

except Exception as e:
    st.error(f"Application Error: {e}")