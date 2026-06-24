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
    # 1. Load Material List: Header is in row 2 (index 1)
    # Force rename of first column to 'PLAID' so the code finds it
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1', engine='python', on_bad_lines='skip')
    mat_df.rename(columns={mat_df.columns[0]: 'PLAID'}, inplace=True)
    mat_df['PLAID'] = mat_df['PLAID'].astype(str).str.strip()
    
    # 2. Load Masterlist
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
    # Masterlist actually has 'PLAID' as a header
    master_df['PLAID'] = master_df['PLAID'].astype(str).str.strip()
    
    # 3. Merge
    df = pd.merge(mat_df, master_df[['PLAID', 'SITE']], on='PLAID', how='left')
    return df.set_index('PLAID')

try:
    material_db = load_data()
    
    # Sidebar
    plaid_list = material_db.index.dropna().tolist()
    selected_plaid = st.sidebar.selectbox("Select PLAID", plaid_list)
    site_name = material_db.loc[selected_plaid, 'SITE']
    st.sidebar.write(f"**Site:** {site_name}")

    if st.button("Generate MRF"):
        site_row = material_db.loc[selected_plaid]
        wb = load_workbook("template_mrf.xlsx")
        ws = wb.active
        
        # Mapping Header Data
        ws['D8'] = f"{selected_plaid} - {site_name}"
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        
        # Mapping Quantities
        # Template scan: rows 16 to 50
        for row in range(16, 50):
            part_no = str(ws[f'A{row}'].value).strip() if ws[f'A{row}'].value else ""
            if part_no in material_db.columns:
                val = site_row[part_no]
                # Only write if value exists
                if pd.notna(val) and val != '':
                    ws[f'D{row}'] = val
        
        output = io.BytesIO()
        wb.save(output)
        
        # Filename per request
        filename = f"NOKIA-FN_06232026-001_{selected_plaid}-{site_name}_MF2_SUBCON.xlsx"
        
        st.success("MRF Generated!")
        st.download_button("📥 Download File", data=output.getvalue(), file_name=filename)

except Exception as e:
    st.error(f"Application Load Error: {e}")