import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from datetime import date
import os

st.set_page_config(page_title="MRF Generator", layout="wide")
st.title("🏗️ Automated MRF Generator with Parts Selection")

# --- DATA LOADING FUNCTIONS ---
@st.cache_data
def load_material_data():
    """Load material database from CSV"""
    mat_df = pd.read_csv("eul_material_list.csv", header=1, encoding='latin1', engine='python', on_bad_lines='skip', index_col=0)
    master_df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
    master_df.rename(columns={master_df.columns[0]: 'PLAID'}, inplace=True)
    master_df.set_index('PLAID', inplace=True)
    return mat_df.join(master_df[['SITE', 'SITE_ADD']], how='left')

@st.cache_data
def load_parts_database():
    """Load parts database from Excel file"""
    try:
        # Check if file exists
        if os.path.exists("PartsDatabase.xlsx"):
            parts_df = pd.read_excel("PartsDatabase.xlsx", sheet_name=0)
            
            # Ensure required columns exist
            required_cols = ['Part_Number', 'Part_Name', 'Category', 'Unit', 'Standard_Qty']
            for col in required_cols:
                if col not in parts_df.columns:
                    # Try to find alternative column names
                    alt_cols = [c for c in parts_df.columns if col.lower() in c.lower()]
                    if alt_cols:
                        parts_df.rename(columns={alt_cols[0]: col}, inplace=True)
                    else:
                        st.warning(f"Column '{col}' not found in PartsDatabase.xlsx. Using default.")
                        parts_df[col] = ""
            
            return parts_df
        else:
            st.warning("PartsDatabase.xlsx not found. Creating sample database...")
            # Create sample database
            sample_parts = pd.DataFrame({
                'Part_Number': ['PART001', 'PART002', 'PART003', 'PART004', 'PART005'],
                'Part_Name': ['Cable Assembly', 'Mounting Bracket', 'Power Supply', 'Antenna', 'Connector Kit'],
                'Category': ['Cables', 'Hardware', 'Electronics', 'Antennas', 'Accessories'],
                'Unit': ['pcs', 'pcs', 'pcs', 'pcs', 'pcs'],
                'Standard_Qty': [10, 5, 3, 2, 15]
            })
            return sample_parts
    except Exception as e:
        st.error(f"Error loading PartsDatabase.xlsx: {e}")
        return pd.DataFrame()

@st.cache_data
def load_site_masterlist():
    """Load site masterlist with ability to add new sites"""
    try:
        if os.path.exists("GLOBE SITE MASTERLIST.csv"):
            df = pd.read_csv("GLOBE SITE MASTERLIST.csv", encoding='latin1', engine='python', on_bad_lines='skip')
            df.rename(columns={df.columns[0]: 'PLAID'}, inplace=True)
            return df
        else:
            # Create empty dataframe with required columns
            return pd.DataFrame(columns=['PLAID', 'SITE', 'SITE_ADD'])
    except Exception as e:
        st.error(f"Error loading site masterlist: {e}")
        return pd.DataFrame()

# --- MAIN APPLICATION ---
# Load all data
material_db = load_material_data()
parts_db = load_parts_database()
site_masterlist = load_site_masterlist()

# --- SIDEBAR - SITE SELECTION ---
st.sidebar.header("Site Selection")

# Get unique site IDs
site_ids = material_db.index.dropna().unique().tolist() if not material_db.empty else []

# Add option to add new site
site_options = site_ids + ["➕ Add New Site"] if site_ids else ["➕ Add New Site"]

selected_plaid = st.sidebar.selectbox("Select Site ID", site_options)

# Handle new site addition
if selected_plaid == "➕ Add New Site":
    st.sidebar.subheader("Add New Site")
    new_plaid = st.sidebar.text_input("New Site ID (PLAID)")
    new_site_name = st.sidebar.text_input("Site Name")
    new_site_address = st.sidebar.text_area("Site Address")
    
    if st.sidebar.button("Save New Site"):
        if new_plaid and new_site_name:
            # Add to material database
            new_row = pd.DataFrame({
                'SITE': [new_site_name],
                'SITE_ADD': [new_site_address]
            }, index=[new_plaid])
            
            # Add columns for all parts with default 0
            for col in material_db.columns:
                if col not in ['SITE', 'SITE_ADD']:
                    new_row[col] = 0
            
            # Add to material_db
            material_db = pd.concat([material_db, new_row])
            st.sidebar.success(f"✅ Site {new_plaid} added successfully!")
            st.rerun()
        else:
            st.sidebar.error("Please enter Site ID and Site Name")

# Get site info if valid selection
if selected_plaid in material_db.index:
    site_info = material_db.loc[selected_plaid]
else:
    site_info = pd.Series({'SITE': '', 'SITE_ADD': ''})

# --- PARTS SELECTION SECTION ---
st.header("🔧 Parts Selection and Quantities")

# Two columns for parts selection
col_parts1, col_parts2 = st.columns(2)

with col_parts1:
    # Category filter
    if not parts_db.empty and 'Category' in parts_db.columns:
        categories = ["All"] + parts_db['Category'].unique().tolist()
        selected_category = st.selectbox("Filter by Category", categories)
        
        if selected_category != "All":
            filtered_parts = parts_db[parts_db['Category'] == selected_category]
        else:
            filtered_parts = parts_db
    else:
        filtered_parts = parts_db

with col_parts2:
    # Search functionality
    search_term = st.text_input("Search Parts", placeholder="Enter part number or name...")
    if search_term:
        filtered_parts = filtered_parts[
            filtered_parts['Part_Number'].str.contains(search_term, case=False, na=False) |
            filtered_parts['Part_Name'].str.contains(search_term, case=False, na=False)
        ]

# Display parts with quantity selection
if not filtered_parts.empty:
    st.subheader("Select Parts and Specify Quantities")
    
    # Create a dataframe for parts selection
    parts_selection = []
    
    # Check if we have existing quantities from material_db
    existing_parts = {}
    if selected_plaid in material_db.index:
        for col in material_db.columns:
            if col not in ['SITE', 'SITE_ADD']:
                val = material_db.loc[selected_plaid, col]
                if pd.notna(val) and str(val).strip() != '' and str(val) != '0':
                    existing_parts[col] = val
    
    # Create selection table
    st.info(f"Found {len(filtered_parts)} parts. Select and set quantities below:")
    
    # Use columns for better layout
    cols = st.columns(4)
    
    with cols[0]:
        st.write("**Part Number**")
    with cols[1]:
        st.write("**Part Name**")
    with cols[2]:
        st.write("**Unit**")
    with cols[3]:
        st.write("**Quantity**")
    
    # Initialize with existing values
    selected_parts = {}
    
    for idx, row in filtered_parts.iterrows():
        part_num = str(row['Part_Number'])
        if pd.isna(part_num) or part_num == 'nan':
            continue
            
        cols = st.columns(4)
        
        with cols[0]:
            st.write(part_num)
        
        with cols[1]:
            part_name = str(row.get('Part_Name', ''))
            st.write(part_name)
        
        with cols[2]:
            unit = str(row.get('Unit', 'pcs'))
            st.write(unit)
        
        with cols[3]:
            default_qty = existing_parts.get(part_num, 0)
            qty = st.number_input(
                f"Qty_{part_num}",
                min_value=0,
                max_value=9999,
                value=int(default_qty) if default_qty > 0 else 0,
                step=1,
                key=f"qty_{part_num}"
            )
            if qty > 0:
                selected_parts[part_num] = qty
    
    # Show summary of selected parts
    if selected_parts:
        st.subheader("📋 Selected Parts Summary")
        summary_df = pd.DataFrame({
            'Part Number': list(selected_parts.keys()),
            'Quantity': list(selected_parts.values())
        })
        st.dataframe(summary_df, use_container_width=True)
        
        # Add option to clear all selections
        if st.button("🗑️ Clear All Selections"):
            for key in st.session_state.keys():
                if key.startswith("qty_"):
                    st.session_state[key] = 0
            st.rerun()
else:
    st.warning("No parts found. Please check your PartsDatabase.xlsx file.")

# --- USER INPUTS ---
st.header("📝 MRF Details")
col1, col2 = st.columns(2)

with col1:
    destination_city = st.text_input("Enter Destination City")
    received_by = st.text_input("Received By")
    mod_site_add = st.text_area(
        "Site Address", 
        value=str(site_info.get('SITE_ADD', '')) if isinstance(site_info, pd.Series) else ''
    )

with col2:
    uploaded_file = st.file_uploader("Upload E-Signature", type=['png', 'jpg', 'jpeg'])

# --- GENERATE MRF BUTTON ---
if st.button("🚀 Generate MRF", type="primary"):
    if not selected_plaid or selected_plaid == "➕ Add New Site":
        st.error("Please select a valid site.")
    elif not destination_city:
        st.error("Please enter Destination City.")
    elif not received_by:
        st.error("Please enter Received By.")
    else:
        try:
            # Load template
            wb = load_workbook("template_mrf.xlsx")
            ws = wb.active
            
            # 1. Replace Text Placeholders
            replacements = {
                "[CITY]": destination_city,
                "[PLAID  - SITE NAME]": f"{selected_plaid} - {site_info.get('SITE', 'Unknown Site')}",
                "[SITE_ADD]": mod_site_add,
                "[RECEIVED BY]": received_by,
                "[DATE_GENERATED]": date.today().strftime("%Y-%m-%d"),
                "[ESIG]": ""
            }
            
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        for placeholder, value in replacements.items():
                            if placeholder in cell.value:
                                cell.value = cell.value.replace(placeholder, value)
            
            # 2. Update Parts Quantities
            current_row = 16
            
            # Clear existing quantities
            for row in range(16, 61):
                ws[f'D{row}'] = ""
            
            # Add selected parts
            for part_num, qty in selected_parts.items():
                if qty > 0:
                    # Check if part exists in template (rows 16-60)
                    found = False
                    for row in range(16, 61):
                        part_cell = ws[f'A{row}']
                        if part_cell.value and str(part_cell.value).strip() == part_num:
                            ws[f'D{row}'] = qty
                            found = True
                            break
                    
                    # If not found, add to new row
                    if not found:
                        current_row += 1
                        ws[f'A{current_row}'] = part_num
                        ws[f'D{current_row}'] = qty
            
            # 3. Add Additional Parts from Material DB (not selected manually)
            if selected_plaid in material_db.index:
                for col in material_db.columns:
                    if col not in ['SITE', 'SITE_ADD'] and col not in selected_parts:
                        val = material_db.loc[selected_plaid, col]
                        if pd.notna(val) and str(val).strip() != '' and str(val) != '0':
                            current_row += 1
                            ws[f'A{current_row}'] = col
                            ws[f'D{current_row}'] = val
            
            # 4. Inject Signature Image
            if uploaded_file:
                try:
                    target_cell = 'D47'
                    
                    # Check if D47 is part of a merge
                    for merged_range in ws.merged_cells.ranges:
                        if 'D47' in merged_range:
                            target_cell = merged_range.start_cell.coordinate
                            break
                    
                    ws[target_cell] = ""
                    
                    img = ExcelImage(uploaded_file)
                    img.width = 120
                    img.height = 60
                    ws.add_image(img, target_cell)
                    
                except Exception as e:
                    st.warning(f"Could not place signature: {e}")
            
            # Save and Trigger Download
            output = io.BytesIO()
            wb.save(output)
            filename = f"NOKIA-MRF_{selected_plaid}-{site_info.get('SITE', 'Unknown')}_{date.today().strftime('%Y%m%d')}.xlsx"
            
            st.success("✅ MRF Generated Successfully!")
            st.download_button(
                "📥 Download Final MRF", 
                data=output.getvalue(), 
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Error generating MRF: {e}")
            st.exception(e)

# --- FOOTER ---
st.divider()
st.caption("ℹ️ Upload PartsDatabase.xlsx with columns: Part_Number, Part_Name, Category, Unit, Standard_Qty")