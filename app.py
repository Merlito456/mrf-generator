import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import date

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

@st.cache_data
def load_data():
    # 1. Load Material List
    # We load without header first, then manually set the columns
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1', engine='python', on_bad_lines='skip')
    # Force the first column to be 'PLAID'
    mat_df.rename(columns={mat_df.columns[0]: 'PLAID'}, inplace=True)
    
    # 2. Load Masterlist
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
    # Ensure 'PLAID' exists (it should, but we force it)
    if 'PLAID' not in master_df.columns:
        master_df.rename(columns={master_df.columns[0]: 'PLAID'}, inplace=True)
        
    # 3. Clean and Merge
    mat_df['PLAID'] = mat_df['PLAID'].astype(str).str.strip()
    master_df['PLAID'] = master_df['PLAID'].astype(str).str.strip()
    
    df = pd.merge(mat_df, master_df[['PLAID', 'SITE']], on='PLAID', how='left')
    return df.set_index('PLAID')

try:
    material_db = load_data()
    
    # UI
    selected_plaid = st.sidebar.selectbox("Select PLAID", material_db.index.tolist())
    site_name = material_db.loc[selected_plaid, 'SITE']
    st.sidebar.write(f"**Site:** {site_name}")

    if st.button("Generate MRF"):
        site_row = material_db.loc[selected_plaid]
        wb = load_workbook("template_mrf.xlsx")
        ws = wb.active
        
        ws['D8'] = f"{selected_plaid} - {site_name}"
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        
        # Map quantities: look through rows 16 to 60
        for row in range(16, 60):
            part_no = str(ws[f'A{row}'].value).strip()
            if part_no in material_db.columns:
                ws[f'D{row}'] = site_row[part_no]
        
        output = io.BytesIO()
        wb.save(output)
        st.download_button("📥 Download MRF", data=output.getvalue(), file_name=f"MRF_{selected_plaid}.xlsx")

except Exception as e:
    st.error(f"Data loading error: {e}")