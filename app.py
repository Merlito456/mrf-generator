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
    # 1. Load Material List: Using header=1 to match your data structure
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1', engine='python', on_bad_lines='skip')
    
    # Rename the first column (Site ID) to 'PLAID'
    mat_df = mat_df.rename(columns={mat_df.columns[0]: 'PLAID'})
    mat_df['PLAID'] = mat_df['PLAID'].astype(str).str.strip()
    
    # 2. Load Masterlist
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
    master_df['PLAID'] = master_df['PLAID'].astype(str).str.strip()
    
    # 3. Merge: Combine Material List and Site Names
    df = pd.merge(mat_df, master_df[['PLAID', 'SITE']], on='PLAID', how='left')
    return df.set_index('PLAID')

try:
    material_db = load_data()

    # --- APP INTERFACE ---
    st.sidebar.header("Site Selection")
    selected_plaid = st.sidebar.selectbox("Select PLAID", material_db.index.tolist())
    
    # Fetch site name from the merged database
    site_name = material_db.loc[selected_plaid, 'SITE']
    st.sidebar.write(f"**Site Name:** {site_name}")

    if st.button("Generate MRF"):
        site_row = material_db.loc[selected_plaid]
        
        # Load Template
        wb = load_workbook("template_mrf.xlsx")
        ws = wb.active
        
        # Fill Metadata (Cells D8, E4)
        ws['D8'] = f"{selected_plaid} - {site_name}"
        ws['E4'] = date.today().strftime("%Y-%m-%d")
        
        # Automated Mapping: Scan rows 16-60 for Part Numbers in Column A
        for row in range(16, 60):
            part_no = str(ws[f'A{row}'].value).strip()
            
            # Check if Part Number exists as a column header in the CSV
            if part_no in material_db.columns:
                ws[f'D{row}'] = site_row[part_no]
        
        # Save to buffer
        output = io.BytesIO()
        wb.save(output)
        
        # Final Naming
        new_filename = f"MRF_{selected_plaid}_{site_name}.xlsx"
        st.success(f"MRF generated for {selected_plaid}!")
        st.download_button("📥 Download MRF File", data=output.getvalue(), file_name=new_filename)

except Exception as e:
    st.error(f"Application error: {e}")
    st.write("Please ensure 'eul_material_list.csv' and 'GLOBE SITE MASTERLIST.csv' are in the root folder.")