import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import date

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

@st.cache_data
def load_data():
    # EUL Material List: First column is PLAID, headers are on row 1 (0-indexed in list)
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1', engine='python', on_bad_lines='skip')
    mat_df.rename(columns={mat_df.columns[0]: 'PLAID'}, inplace=True)
    
    # Masterlist
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
    
    # Merge and Clean
    mat_df['PLAID'] = mat_df['PLAID'].astype(str).str.strip()
    master_df['PLAID'] = master_df['PLAID'].astype(str).str.strip()
    df = pd.merge(mat_df, master_df[['PLAID', 'SITE']], on='PLAID', how='left')
    return df.set_index('PLAID')

try:
    material_db = load_data()
    selected_plaid = st.sidebar.selectbox("Select PLAID", material_db.index.tolist())
    site_name = material_db.loc[selected_plaid, 'SITE']
    st.sidebar.write(f"**Site:** {site_name}")

    if st.button("Generate MRF"):
        site_row = material_db.loc[selected_plaid]
        wb = load_workbook("template_mrf.xlsx")
        ws = wb.active
        
        # 1. Update Metadata
        ws['D8'] = f"{selected_plaid} - {site_name}"
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        
        # 2. Precise Mapping: Iterate through rows 16 to 50
        # We look at Column A for Part Number, write to Column D (REQ)
        for row in range(16, 50):
            part_no = str(ws[f'A{row}'].value).strip() if ws[f'A{row}'].value else ""
            
            # Match part_no against CSV headers
            if part_no in material_db.columns:
                val = site_row[part_no]
                if pd.notna(val):
                    ws[f'D{row}'] = val
        
        output = io.BytesIO()
        wb.save(output)
        
        # 3. Required Filename Format
        filename = f"NOKIA-FN_06232026-001_{selected_plaid}-{site_name}_MF2_SUBCON.xlsx"
        
        st.success(f"MRF generated successfully!")
        st.download_button("📥 Download MRF File", data=output.getvalue(), file_name=filename)

except Exception as e:
    st.error(f"Error: {e}")