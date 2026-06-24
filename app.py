import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import date

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

# 1. Load Data with header=1 to handle your CSV structure
@st.cache_data
def load_data():
    # Load Material List (using header=1 because your headers are in the 2nd row)
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1')
    # Rename the first column (Site ID) to 'PLAID' for easier lookup
    mat_df = mat_df.rename(columns={mat_df.columns[0]: 'PLAID'})
    
    # Load Masterlist to get Site Names
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1')
    
    # Merge them so we have PLAID and Site Name together
    df = pd.merge(mat_df, master_df[['PLAID', 'SITE']], on='PLAID', how='left')
    return df.set_index('PLAID')

material_db = load_data()

# 2. Site Selection
selected_plaid = st.sidebar.selectbox("Select PLAID", material_db.index.tolist())
# Fetch site name automatically from our merged database
site_name = material_db.loc[selected_plaid, 'SITE']
st.sidebar.write(f"**Selected Site:** {site_name}")

if st.button("Generate MRF"):
    site_row = material_db.loc[selected_plaid]
    wb = load_workbook("template_mrf.xlsx")
    ws = wb.active
    
    # Fill Metadata
    ws['D8'] = f"{selected_plaid} - {site_name}"
    ws['E4'] = date.today().strftime("%Y-%m-%d")
    
    # Map Quantities
    for row in range(16, 60):
        part_no = str(ws[f'A{row}'].value).strip()
        if part_no in material_db.columns:
            ws[f'D{row}'] = site_row[part_no]
    
    output = io.BytesIO()
    wb.save(output)
    
    new_filename = f"MRF_{selected_plaid}_{site_name}.xlsx"
    st.download_button("📥 Download MRF", data=output.getvalue(), file_name=new_filename)