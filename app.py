import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from datetime import date

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

# --- DATA LOADING ---
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

# --- USER INPUTS ---
col1, col2 = st.columns(2)
with col1:
    destination_city = st.text_input("Enter Destination City")
    received_by = st.text_input("Received By")
    mod_site_add = st.text_area("Site Address", value=str(site_info.get('SITE_ADD', '')))
with col2:
    uploaded_file = st.file_uploader("Upload E-Signature", type=['png', 'jpg'])

if st.button("Generate MRF"):
    wb = load_workbook("template_mrf.xlsx")
    ws = wb.active
    
# 1. Replace Text Placeholders
    replacements = {
        "[CITY]": destination_city,
        "[PLAID  - SITE NAME]": f"{selected_plaid} - {site_info['SITE']}",
        "[SITE_ADD]": mod_site_add,
        "[RECEIVED BY]": received_by,
        "[DATE_GENERATED]": date.today().strftime("%Y-%m-%d"),
        "[ESIG]": "" # We clear the tag here so it doesn't leave stray brackets
    }
    
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for placeholder, value in replacements.items():
                    if placeholder in cell.value:
                        # This replaces ONLY the tag, not the whole cell text
                        cell.value = cell.value.replace(placeholder, value)

# 2. Map Quantities & Add Missing Items
    current_row = 16
    processed_parts = []
    
    # First: Update existing parts in rows 16-60
    for row in range(16, 61):
        part_cell = ws[f'A{row}']
        part_no = str(part_cell.value).strip() if part_cell.value else ""
        
        if part_no in material_db.columns:
            val = site_info[part_no]
            # Only write if there is a value; otherwise, leave blank
            if pd.notna(val) and str(val).strip() != '' and str(val) != '0':
                ws[f'D{row}'] = val
            else:
                ws[f'D{row}'] = "" # Explicitly clear if no requirement
            processed_parts.append(part_no)
        
        if row > current_row: current_row = row

    # Second: Add missing items that are in DB but not in template
    for part_col in material_db.columns:
        if part_col not in processed_parts and part_col not in ['SITE', 'SITE_ADD']:
            val = site_info[part_col]
            if pd.notna(val) and str(val).strip() != '' and str(val) != '0':
                current_row += 1
                ws[f'A{current_row}'] = part_col # Part Number
                ws[f'D{current_row}'] = val      # Quantity
    # 3. Inject Signature Image (Targeting D47 to avoid Merge Errors)
    if uploaded_file:
        try:
            target_cell = 'D47'
            
            # Check if D47 is part of a merge
            for merged_range in ws.merged_cells.ranges:
                if 'D47' in merged_range:
                    # Use the top-left cell of the merge as the target
                    target_cell = merged_range.start_cell.coordinate
                    break
            
            # Clear text in the top-left cell
            ws[target_cell] = ""
            
            img = ExcelImage(uploaded_file)
            img.width = 120
            img.height = 60
            ws.add_image(img, target_cell)
            
        except Exception as e:
            st.error(f"Error placing signature: {e}")
    # Save and Trigger Download
    output = io.BytesIO()
    wb.save(output)
    filename = f"NOKIA-FN_06232026-001_{selected_plaid}-{site_info['SITE']}_MF2_SUBCON.xlsx"
    st.success("MRF Generated!")
    st.download_button("📥 Download Final MRF", data=output.getvalue(), file_name=filename)