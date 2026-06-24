import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import date

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

# --- DATA LOADING ---
@st.cache_data
def load_data():
    # 1. Load Material List: Header at row 2 (index 1), first column is Site ID
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1', engine='python', on_bad_lines='skip', index_col=0)
    
    # 2. Load Masterlist: First column is PLAID/Site ID
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
    master_df.rename(columns={master_df.columns[0]: 'PLAID'}, inplace=True)
    master_df.set_index('PLAID', inplace=True)
    
    # 3. Merge to get Site Name
    df = mat_df.join(master_df[['SITE']], how='left')
    return df

try:
    material_db = load_data()
    
    # Sidebar: Select Site
    selected_plaid = st.sidebar.selectbox("Select Site ID", material_db.index.dropna().unique().tolist())
    site_info = material_db.loc[selected_plaid]
    site_name = site_info['SITE'] if 'SITE' in site_info else "Unknown"

    if st.button("Generate MRF"):
        # Load the existing template
        wb = load_workbook("template_mrf.xlsx")
        ws = wb.active
        
        # --- PRECISE MAPPING (Preserves formatting) ---
        # Destination Code value goes into D8
        ws['D8'] = f"{selected_plaid} - {site_name}"
        
        # Site ID value goes into D10
        ws['D10'] = selected_plaid
        
        # Date value goes into E4
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        
        # --- QUANTITY MAPPING ---
        # Scan rows 16 to 60, Column A (Part #) -> Column D (Qty)
        for row in range(16, 60):
            part_cell = ws[f'A{row}']
            part_no = str(part_cell.value).strip() if part_cell.value else ""
            
            # If the part is in our DB, write the quantity to Column D
            if part_no in material_db.columns:
                val = site_info[part_no]
                if pd.notna(val) and str(val).strip() != '':
                    ws[f'D{row}'] = val
        
        # Save output to buffer
        output = io.BytesIO()
        wb.save(output)
        
        # Filename per request
        filename = f"NOKIA-FN_06232026-001_{selected_plaid}-{site_name}_MF2_SUBCON.xlsx"
        
        st.success("MRF Generated!")
        st.download_button("📥 Download File", data=output.getvalue(), file_name=filename)

except Exception as e:
    st.error(f"Application Error: {e}")