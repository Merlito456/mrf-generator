import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook
from datetime import date

# --- CONFIGURATION ---
DB_FILE = "eul_material_list.csv"
TEMPLATE_FILE = "template_mrf.xlsx"

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator")

# --- DATA LOADING ---
@st.cache_data
def load_db():
    try:
        # We skip potential bad lines and force the engine to handle uneven columns
        return pd.read_csv(
            DB_FILE, 
            encoding='latin1', 
            engine='python', 
            on_bad_lines='skip'
        ).set_index('PLAID')
    except Exception as e:
        st.error(f"Error loading database: {e}")
        return None

material_db = load_db()

# --- APP INTERFACE ---
if material_db is not None:
    st.sidebar.header("Site Selection")
    # Clean the index to remove potential whitespace
    plaid_list = [str(i).strip() for i in material_db.index.tolist()]
    selected_plaid = st.sidebar.selectbox("Select PLAID", plaid_list)
    site_name = st.sidebar.text_input("Enter Site Name for Output")

    if st.button("Generate MRF"):
        try:
            site_row = material_db.loc[selected_plaid]
            wb = load_workbook(TEMPLATE_FILE)
            ws = wb.active
            
            # Update Metadata
            ws['D8'] = f"{selected_plaid} - {site_name}"
            ws['E4'] = date.today().strftime("%Y-%m-%d")
            
            # --- AUTOMATED MAPPING ---
            # Part Numbers are in Column A (1), Quantities in Column D (4)
            for row in range(16, 60):
                part_no = str(ws[f'A{row}'].value).strip()
                
                # If part_no found in CSV database columns, update quantity
                if part_no in material_db.columns:
                    val = site_row[part_no]
                    ws[f'D{row}'] = val
            
            # Save to buffer
            output = io.BytesIO()
            wb.save(output)
            
            new_filename = TEMPLATE_FILE.replace("PLAID-SITE NAME", f"{selected_plaid}-{site_name}")
            st.success(f"MRF generated for {selected_plaid}!")
            st.download_button("📥 Download MRF File", data=output.getvalue(), file_name=new_filename)
            
        except Exception as e:
            st.error(f"Processing error: {e}")