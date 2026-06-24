import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from datetime import date

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

@st.cache_data
def load_data():
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1', engine='python', on_bad_lines='skip', index_col=0)
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
    master_df.rename(columns={master_df.columns[0]: 'PLAID'}, inplace=True)
    master_df.set_index('PLAID', inplace=True)
    return mat_df.join(master_df[['SITE', 'SITE_ADD']], how='left')

material_db = load_data()
selected_plaid = st.sidebar.selectbox("Select Site ID", material_db.index.dropna().unique().tolist())
site_info = material_db.loc[selected_plaid]

# --- MODIFIABLE INPUTS ---
col1, col2 = st.columns(2)
with col1:
    destination_city = st.text_input("Enter Destination City")
    received_by = st.text_input("Received By")
    # Editable Site Address: Starts with data from masterlist, user can overwrite
    mod_site_add = st.text_area("Site Address", value=str(site_info.get('SITE_ADD', '')))
with col2:
    uploaded_file = st.file_uploader("Upload E-Signature", type=['png', 'jpg'])

if st.button("Generate MRF"):
    wb = load_workbook("template_mrf.xlsx")
    ws = wb.active
    
    # 1. Replace Placeholders
    replacements = {
        "[CITY]": destination_city,
        "[PLAID  - SITE NAME]": f"{selected_plaid} - {site_info['SITE']}",
        "[SITE_ADD]": mod_site_add,
        "[RECEIVED BY]": received_by,
        "[DATE_GENERATED]": date.today().strftime("%Y-%m-%d")
    }
    
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for placeholder, value in replacements.items():
                    if placeholder in cell.value:
                        cell.value = cell.value.replace(placeholder, value)

    # 2. Map Quantities (Rows 16-60, Col A to D)
    for row in range(16, 60):
        part_no = str(ws[f'A{row}'].value).strip() if ws[f'A{row}'].value else ""
        if part_no in material_db.columns:
            val = site_info[part_no]
            if pd.notna(val) and str(val).strip() != '':
                ws[f'D{row}'] = val

    # 3. Add Signature
    if uploaded_file:
        img = ExcelImage(uploaded_file)
        img.width = 120
        img.height = 60
        ws.add_image(img, 'E20') # Ensure this matches your template's [ESIG] location

    output = io.BytesIO()
    wb.save(output)
    st.download_button("📥 Download Final MRF", data=output.getvalue(), file_name=f"MRF_{selected_plaid}.xlsx")