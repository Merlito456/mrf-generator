import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import date

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

@st.cache_data
def load_data():
    # 1. Load Material List: Header is row 2. 
    # Use index_col=0 to force the first column (Site ID) to be the index.
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1', engine='python', on_bad_lines='skip', index_col=0)
    
    # 2. Load Masterlist: First column is PLAID.
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
    
    # 3. Standardize and Merge
    # Rename master_df's first column to match the index of mat_df
    master_df.rename(columns={master_df.columns[0]: 'PLAID'}, inplace=True)
    master_df.set_index('PLAID', inplace=True)
    
    # Merge on the index
    df = mat_df.join(master_df[['SITE']], how='left')
    return df

try:
    material_db = load_data()
    
    # Sidebar Selection
    selected_plaid = st.sidebar.selectbox("Select PLAID (Site ID)", material_db.index.dropna().unique().tolist())
    site_name = material_db.loc[selected_plaid, 'SITE']
    st.sidebar.write(f"**Site Name:** {site_name}")

    if st.button("Generate MRF"):
        site_row = material_db.loc[selected_plaid]
        wb = load_workbook("template_mrf.xlsx")
        ws = wb.active
        
        # Fill Metadata
        ws['D8'] = f"{selected_plaid} - {site_name}"
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        
        # Mapping: Rows 16 to 50, Column A (Part #) -> Column D (Qty)
        for row in range(16, 50):
            part_no = str(ws[f'A{row}'].value).strip() if ws[f'A{row}'].value else ""
            
            # Match part_no against CSV column headers
            if part_no in material_db.columns:
                val = site_row[part_no]
                if pd.notna(val) and val != '':
                    ws[f'D{row}'] = val
        
        output = io.BytesIO()
        wb.save(output)
        
        # Requested Format
        filename = f"NOKIA-FN_06232026-001_{selected_plaid}-{site_name}_MF2_SUBCON.xlsx"
        
        st.success("MRF Generated!")
        st.download_button("📥 Download File", data=output.getvalue(), file_name=filename)

except Exception as e:
    st.error(f"Data loading error: {e}")