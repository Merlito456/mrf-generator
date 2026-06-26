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
    """Load parts database from Excel file with PART NUMBER and DESCRIPTION columns"""
    try:
        # Check if file exists
        if os.path.exists("PartsDatabase.xlsx"):
            parts_df = pd.read_excel("PartsDatabase.xlsx", sheet_name=0)
            
            # Map columns to expected format
            column_mapping = {}
            
            # Look for PART NUMBER column
            part_cols = [col for col in parts_df.columns if 'PART' in col.upper() and 'NUMBER' in col.upper()]
            if part_cols:
                column_mapping[part_cols[0]] = 'Part_Number'
            elif 'PART NUMBER' in parts_df.columns:
                column_mapping['PART NUMBER'] = 'Part_Number'
            elif 'PART_NO' in parts_df.columns:
                column_mapping['PART_NO'] = 'Part_Number'
            elif 'PART_NUMBER' in parts_df.columns:
                column_mapping['PART_NUMBER'] = 'Part_Number'
            elif 'PART' in parts_df.columns:
                column_mapping['PART'] = 'Part_Number'
            
            # Look for DESCRIPTION column
            desc_cols = [col for col in parts_df.columns if 'DESCRIPTION' in col.upper() or 'DESC' in col.upper()]
            if desc_cols:
                column_mapping[desc_cols[0]] = 'Part_Name'
            elif 'DESCRIPTION' in parts_df.columns:
                column_mapping['DESCRIPTION'] = 'Part_Name'
            elif 'DESC' in parts_df.columns:
                column_mapping['DESC'] = 'Part_Name'
            elif 'PART_NAME' in parts_df.columns:
                column_mapping['PART_NAME'] = 'Part_Name'
            
            # Rename columns if mapping found
            if column_mapping:
                parts_df.rename(columns=column_mapping, inplace=True)
            
            # Add default columns if missing
            if 'Category' not in parts_df.columns:
                parts_df['Category'] = 'General'
            if 'Unit' not in parts_df.columns:
                parts_df['Unit'] = 'pcs'
            if 'Standard_Qty' not in parts_df.columns:
                parts_df['Standard_Qty'] = 0
            
            # Clean up - remove rows with empty part numbers
            parts_df = parts_df[parts_df['Part_Number'].notna()]
            parts_df['Part_Number'] = parts_df['Part_Number'].astype(str).str.strip()
            parts_df = parts_df[parts_df['Part_Number'] != '']
            
            return parts_df
        else:
            st.warning("PartsDatabase.xlsx not found. Please ensure it exists in the current directory.")
            return pd.DataFrame()
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
st.sidebar.header("🏢 Site Selection")

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
    
    if st.sidebar.button("💾 Save New Site"):
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

# Display parts database info
if not parts_db.empty:
    st.info(f"📊 Loaded {len(parts_db)} parts from PartsDatabase.xlsx")
    
    # Preview first few rows
    with st.expander("📋 Preview Parts Database"):
        st.dataframe(parts_db[['Part_Number', 'Part_Name', 'Category']].head(10), use_container_width=True)
        if len(parts_db) > 10:
            st.caption(f"... and {len(parts_db) - 10} more parts")
else:
    st.error("❌ PartsDatabase.xlsx not loaded. Please check the file exists and has 'PART NUMBER' and 'DESCRIPTION' columns.")

# Two columns for parts selection
col_parts1, col_parts2 = st.columns(2)

with col_parts1:
    # Category filter - create categories from parts if available
    if not parts_db.empty and 'Category' in parts_db.columns:
        categories = ["All"] + sorted(parts_db['Category'].dropna().unique().tolist())
        selected_category = st.selectbox("🔍 Filter by Category", categories)
        
        if selected_category != "All":
            filtered_parts = parts_db[parts_db['Category'] == selected_category]
        else:
            filtered_parts = parts_db
    else:
        filtered_parts = parts_db

with col_parts2:
    # Search functionality
    search_term = st.text_input("🔎 Search Parts", placeholder="Enter part number or description...")
    if search_term and not filtered_parts.empty:
        filtered_parts = filtered_parts[
            filtered_parts['Part_Number'].astype(str).str.contains(search_term, case=False, na=False) |
            filtered_parts['Part_Name'].astype(str).str.contains(search_term, case=False, na=False)
        ]

# Display parts with quantity selection
if not filtered_parts.empty:
    st.subheader("📦 Select Parts and Specify Quantities")
    
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
    
    # Create selection table with better layout
    st.markdown("### Select parts and enter quantities")
    
    # Use columns for better layout
    col_labels = st.columns([2, 3, 1, 1.5])
    with col_labels[0]:
        st.write("**Part Number**")
    with col_labels[1]:
        st.write("**Description**")
    with col_labels[2]:
        st.write("**Unit**")
    with col_labels[3]:
        st.write("**Quantity**")
    
    st.divider()
    
    # Initialize with existing values
    selected_parts = {}
    
    for idx, row in filtered_parts.iterrows():
        part_num = str(row['Part_Number'])
        if pd.isna(part_num) or part_num == 'nan' or part_num == '':
            continue
            
        cols = st.columns([2, 3, 1, 1.5])
        
        with cols[0]:
            st.code(part_num, language=None)
        
        with cols[1]:
            part_name = str(row.get('Part_Name', ''))
            st.write(part_name)
        
        with cols[2]:
            unit = str(row.get('Unit', 'pcs'))
            st.write(unit)
        
        with cols[3]:
            default_qty = existing_parts.get(part_num, 0)
            # If default quantity exists, use it; otherwise use Standard_Qty
            if default_qty == 0 and 'Standard_Qty' in row:
                default_qty = int(row['Standard_Qty']) if pd.notna(row['Standard_Qty']) else 0
            
            qty = st.number_input(
                f"Qty_{part_num}",
                min_value=0,
                max_value=9999,
                value=int(default_qty) if default_qty > 0 else 0,
                step=1,
                key=f"qty_{part_num}",
                label_visibility="collapsed"
            )
            if qty > 0:
                selected_parts[part_num] = qty
        
        st.divider()
    
    # Show summary of selected parts
    if selected_parts:
        st.subheader("📋 Selected Parts Summary")
        summary_data = []
        for part_num, qty in selected_parts.items():
            part_row = filtered_parts[filtered_parts['Part_Number'] == part_num].iloc[0] if not filtered_parts[filtered_parts['Part_Number'] == part_num].empty else None
            if part_row is not None:
                summary_data.append({
                    'Part Number': part_num,
                    'Description': part_row.get('Part_Name', ''),
                    'Quantity': qty,
                    'Unit': part_row.get('Unit', 'pcs')
                })
        
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            
            # Quick actions
            col_actions1, col_actions2 = st.columns(2)
            with col_actions1:
                total_parts = len(selected_parts)
                total_qty = sum(selected_parts.values())
                st.info(f"📦 Total: {total_parts} part types, {total_qty} total quantity")
            
            with col_actions2:
                if st.button("🗑️ Clear All Selections", use_container_width=True):
                    for key in st.session_state.keys():
                        if key.startswith("qty_"):
                            st.session_state[key] = 0
                    st.rerun()
else:
    if not parts_db.empty:
        st.warning("No parts match your filters. Try adjusting your search or category filter.")
    else:
        st.warning("⚠️ No parts loaded. Please check PartsDatabase.xlsx file.")

# --- USER INPUTS ---
st.header("📝 MRF Details")
col1, col2 = st.columns(2)

with col1:
    destination_city = st.text_input("📍 Enter Destination City")
    received_by = st.text_input("👤 Received By")
    mod_site_add = st.text_area(
        "🏠 Site Address", 
        value=str(site_info.get('SITE_ADD', '')) if isinstance(site_info, pd.Series) else ''
    )

with col2:
    uploaded_file = st.file_uploader("✍️ Upload E-Signature", type=['png', 'jpg', 'jpeg'])
    if uploaded_file:
        st.image(uploaded_file, caption="Signature Preview", width=200)

# --- GENERATE MRF BUTTON ---
st.divider()
col_gen1, col_gen2, col_gen3 = st.columns([1, 2, 1])
with col_gen2:
    generate_button = st.button("🚀 Generate MRF", type="primary", use_container_width=True)

if generate_button:
    if not selected_plaid or selected_plaid == "➕ Add New Site":
        st.error("❌ Please select a valid site.")
    elif not destination_city:
        st.error("❌ Please enter Destination City.")
    elif not received_by:
        st.error("❌ Please enter Received By.")
    elif not selected_parts:
        st.warning("⚠️ No parts selected. Please select at least one part with quantity > 0.")
    else:
        try:
            with st.spinner("Generating MRF..."):
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
                
                # Clear existing quantities in rows 16-60
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
                        st.warning(f"⚠️ Could not place signature: {e}")
                
                # Save and Trigger Download
                output = io.BytesIO()
                wb.save(output)
                filename = f"NOKIA-MRF_{selected_plaid}-{site_info.get('SITE', 'Unknown')}_{date.today().strftime('%Y%m%d')}.xlsx"
                
                st.success("✅ MRF Generated Successfully!")
                st.balloons()
                st.download_button(
                    "📥 Download Final MRF", 
                    data=output.getvalue(), 
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"❌ Error generating MRF: {e}")
            st.exception(e)

# --- FOOTER ---
st.divider()
st.caption("ℹ️ **PartsDatabase.xlsx** should have columns: **PART NUMBER** and **DESCRIPTION** (case-insensitive)")
st.caption("📌 Additional columns like Category, Unit, Standard_Qty are optional but recommended for better organization.")