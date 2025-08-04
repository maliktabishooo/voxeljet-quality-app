import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
from io import BytesIO
import datetime
import re
import os
import time

# Enable automatic reload
def autoreload():
    st.runtime.legacy_caching.clear_cache()
    st.experimental_rerun()

# Check for file changes and reload
def watch_files():
    watched_files = ['x_measurement.png', 'y_measurement.png', 'z_measurement.png', 'brafe_logo.png']
    file_stats = {f: os.stat(f).st_mtime if os.path.exists(f) else 0 for f in watched_files}
    
    while True:
        time.sleep(1)
        for f in watched_files:
            if os.path.exists(f):
                current_mtime = os.stat(f).st_mtime
                if current_mtime > file_stats[f]:
                    file_stats[f] = current_mtime
                    autoreload()

# Start file watcher in a separate thread
if not st._is_running_with_streamlit:
    import threading
    threading.Thread(target=watch_files, daemon=True).start()

# Load measurement images and Brafe logo with error handling
@st.cache_resource
def load_images():
    images = {}
    try:
        images['x_img'] = Image.open('x_measurement.png')
    except FileNotFoundError:
        images['x_img'] = None
        st.warning("Image 'x_measurement.png' not found. Using placeholder.")
    try:
        images['y_img'] = Image.open('y_measurement.png')
    except FileNotFoundError:
        images['y_img'] = None
        st.warning("Image 'y_measurement.png' not found. Using placeholder.")
    try:
        images['z_img'] = Image.open('z_measurement.png')
    except FileNotFoundError:
        images['z_img'] = None
        st.warning("Image 'z_measurement.png' not found. Using placeholder.")
    try:
        images['brafe_logo'] = Image.open('brafe_logo.png')
    except FileNotFoundError:
        images['brafe_logo'] = None
        st.warning("Image 'brafe_logo.png' not found. Using placeholder.")
    return images

images = load_images()

# App configuration with updated blue theme
st.set_page_config(
    page_title="Brafe Engineering Quality Control",
    page_icon=images['brafe_logo'] if images['brafe_logo'] else ":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for blue Brafe theme
st.markdown(
    """
    <style>
    .stApp {
        background-color: #e6f0f9;
        color: #003366;
    }
    .stHeader {
        background-color: #003366;
        padding: 15px;
        border-radius: 5px;
        color: white;
    }
    .stTabs [data-baseweb="tab-list"] {
        background-color: #cce0f5;
        border-radius: 5px;
    }
    .stTabs [data-baseweb="tab"] {
        color: #003366;
        background-color: #cce0f5;
        transition: background-color 0.3s;
        font-weight: bold;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #a3c6f0;
    }
    .stTabs [data-baseweb="tab--selected"] {
        background-color: #00509d;
        color: white;
    }
    .fail-metric {
        background-color: #ffcccc;
        padding: 5px;
        border-radius: 3px;
        animation: pulse 1s infinite;
        color: #cc0000;
    }
    .pass-metric {
        background-color: #d4edda;
        padding: 5px;
        border-radius: 3px;
        color: #155724;
    }
    @keyframes pulse {
        0% { transform: scale(1); }
        50% { transform: scale(1.05); }
        100% { transform: scale(1); }
    }
    .excel-header {
        background-color: #003366;
        color: white;
        font-weight: bold;
    }
    .sidebar-section {
        padding: 10px;
        margin-bottom: 15px;
        border-radius: 5px;
        background-color: #f0f7ff;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Initialize session state
def initialize_session_state():
    if 'operator_name' not in st.session_state:
        st.session_state.operator_name = "Your Name"
    if 'test_id' not in st.session_state:
        st.session_state.test_id = "TEST-001"
    if 'file_dimensions' not in st.session_state:
        st.session_state.file_dimensions = {}
    if 'loi_results' not in st.session_state:
        st.session_state.loi_results = None

initialize_session_state()

# Sidebar with Brafe branding
with st.sidebar:
    # Brafe logo
    if images['brafe_logo']:
        st.image(images['brafe_logo'], width=150)
    else:
        st.markdown("**Brafe Logo Placeholder**")
    
    # Test Information
    st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
    st.subheader("Test Information")
    st.session_state.operator_name = st.text_input("Operator Name", st.session_state.operator_name, key="global_operator")
    st.session_state.test_id = st.text_input("Test ID", st.session_state.test_id, key="global_test_id")
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Engineering Resources
    st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
    st.header("Brafe Engineering Resources")
    st.markdown("Quality Control - LOI and 3 Point Bend Test methods")
    st.markdown("[Technical Support](mailto:info@brafe.com)")
    st.markdown("Hotline: +44 (0) 1394 380 000")
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Test Specifications
    st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
    st.divider()
    st.caption("Test Specifications")
    st.metric("Nominal Bending Strength", "260 N/cm²")
    st.metric("Test Bar Dimensions", "172 × 22.4 × 22.4 mm")
    st.metric("LOI Sample Weight", "30 g")
    st.metric("Dimensional Tolerance", "±0.45 mm")
    st.markdown('</div>', unsafe_allow_html=True)

# Main app
st.markdown('<div class="stHeader"><h1>Brafe Engineering Quality Control Dashboard</h1></div>', unsafe_allow_html=True)
st.subheader("PDB Process Quality Inspection for Printed Parts")

tab1, tab2, tab3 = st.tabs(["Dimensional Check", "3-Point Bend Test", "Loss on Ignition (LOI)"])

with tab1:
    st.header("Dimensional Measurement Verification")
    st.caption("Verify test bar dimensions according to section 3.3 of Quality Control Manual")
    
    with st.expander("📏 Measurement Instructions", expanded=True):
        cols = st.columns(3)
        with cols[0]:
            if images['x_img']:
                st.image(images['x_img'], caption="Measure Dimension X (Length)", use_container_width=True)
            else:
                st.markdown("**Image Placeholder: Measure Dimension X (Length)**")
            st.info("**X-Dimension:**\n- Length direction\n- Nominal: 172 mm")
        with cols[1]:
            if images['y_img']:
                st.image(images['y_img'], caption="Measure Dimension Y (Width)", use_container_width=True)
            else:
                st.markdown("**Image Placeholder: Measure Dimension Y (Width)**")
            st.info("**Y-Dimension:**\n- Width direction\n- Nominal: 22.4 mm")
        with cols[2]:
            if images['z_img']:
                st.image(images['z_img'], caption="Measure Dimension Z (Height)", use_container_width=True)
            else:
                st.markdown("**Image Placeholder: Measure Dimension Z (Height)**")
            st.info("**Z-Dimension:**\n- Height direction\n- Nominal: 22.4 mm")
        
        st.markdown("""
        **Procedure:**
        1. Ensure test bar has rested in sand for ≥6 hours
        2. Clean loose sand from test bar
        3. Measure each dimension with measuring slide
        4. Compare with nominal values (tolerance ±0.45mm)
        """)
    
    st.divider()
    
    with st.form("dimensional_check"):
        st.subheader("Enter Measured Dimensions (mm)")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            x_measured = st.number_input("X-Dimension (Length)", 
                                         min_value=150.0,
                                         max_value=190.0,
                                         value=172.0,
                                         step=0.1,
                                         format="%.1f")
        with col2:
            y_measured = st.number_input("Y-Dimension (Width)", 
                                         min_value=15.0,
                                         max_value=30.0,
                                         value=22.4,
                                         step=0.1,
                                         format="%.1f")
        with col3:
            z_measured = st.number_input("Z-Dimension (Height)", 
                                         min_value=15.0,
                                         max_value=30.0,
                                         value=22.4,
                                         step=0.1,
                                         format="%.1f")
        
        submitted = st.form_submit_button("Verify Dimensions")
        
        if submitted:
            nominal_x, nominal_y, nominal_z = 172.0, 22.4, 22.4
            tolerance = 0.45
            
            x_deviation = x_measured - nominal_x
            y_deviation = y_measured - nominal_y
            z_deviation = z_measured - nominal_z
            
            x_status = "✅ Pass" if abs(x_deviation) <= tolerance else "❌ Fail"
            y_status = "✅ Pass" if abs(y_deviation) <= tolerance else "❌ Fail"
            z_status = "✅ Pass" if abs(z_deviation) <= tolerance else "❌ Fail"
            
            st.subheader("Verification Results")
            cols = st.columns(3)
            with cols[0]:
                st.metric("X-Dimension", f"{x_measured:.1f} mm", 
                          delta=f"{x_deviation:.1f} mm",
                          delta_color="normal" if x_status == "✅ Pass" else "inverse")
                st.markdown(f'<div class="{"pass-metric" if x_status == "✅ Pass" else "fail-metric"}">{x_status}</div>', 
                            unsafe_allow_html=True)
            with cols[1]:
                st.metric("Y-Dimension", f"{y_measured:.1f} mm", 
                          delta=f"{y_deviation:.1f} mm",
                          delta_color="normal" if y_status == "✅ Pass" else "inverse")
                st.markdown(f'<div class="{"pass-metric" if y_status == "✅ Pass" else "fail-metric"}">{y_status}</div>', 
                            unsafe_allow_html=True)
            with cols[2]:
                st.metric("Z-Dimension", f"{z_measured:.1f} mm", 
                          delta=f"{z_deviation:.1f} mm",
                          delta_color="normal" if z_status == "✅ Pass" else "inverse")
                st.markdown(f'<div class="{"pass-metric" if z_status == "✅ Pass" else "fail-metric"}">{z_status}</div>', 
                            unsafe_allow_html=True)
            
            # Add total dimension metric
            total_dimension = x_measured + y_measured + z_measured
            st.markdown("---")
            st.metric("Total Measured Dimensions", f"{total_dimension:.1f} mm")
            
            if all(status == "✅ Pass" for status in [x_status, y_status, z_status]):
                st.success("All dimensions within specification!")
            else:
                st.error("Some dimensions out of tolerance. Check print parameters.")
                st.markdown("""
                **Troubleshooting Tips:**
                - Check offset values in defr3d.ini file
                - Verify zCompensation parameter
                - Ensure proper printer calibration
                """)

with tab2:
    st.header("3-Point Bend Test Analysis")
    st.caption("Calculate bending strength according to section 3.4 of Quality Control Manual")
    
    # Global parameters
    nominal_strength = 260  # N/cm²
    st.markdown(f"**Nominal Bending Strength:** {nominal_strength} N/cm²")
    
    bend_files = st.file_uploader("Upload CSV files from bend test machine",
                                  type=["csv"],
                                  accept_multiple_files=True,
                                  help="Should contain force measurements in last column (Newtons)")
    
    if bend_files:
        # Initialize lists to store results
        results = []
        dfs = []
        dimension_entries = []
        
        # Create dimension input section
        with st.expander("⚙️ Set Test Bar Dimensions for Each File", expanded=True):
            st.subheader("Enter Dimensions for Each Test Bar (mm)")
            
            # Create columns for headers
            header_cols = st.columns([3, 2, 2, 2])
            header_cols[0].markdown("**Filename**")
            header_cols[1].markdown("**Support Span (L)**")
            header_cols[2].markdown("**Width (b)**")
            header_cols[3].markdown("**Height (h)**")
            
            # Create input rows for each file
            for i, bend_file in enumerate(bend_files):
                filename = bend_file.name
                cols = st.columns([3, 2, 2, 2])
                
                # Filename display
                cols[0].markdown(f"`{filename}`")
                
                # Dimension inputs with saved state
                key_prefix = f"dim_{i}"
                
                # Try to get saved dimensions or use defaults
                default_L = st.session_state.file_dimensions.get(f"{filename}_L", 172.0)
                default_b = st.session_state.file_dimensions.get(f"{filename}_b", 22.4)
                default_h = st.session_state.file_dimensions.get(f"{filename}_h", 22.4)
                
                L = cols[1].number_input("L", 
                                         min_value=1.0,
                                         value=default_L,
                                         step=0.1,
                                         format="%.1f",
                                         key=f"{key_prefix}_L",
                                         label_visibility="collapsed")
                b = cols[2].number_input("b", 
                                         min_value=1.0,
                                         value=default_b,
                                         step=0.1,
                                         format="%.1f",
                                         key=f"{key_prefix}_b",
                                         label_visibility="collapsed")
                h = cols[3].number_input("h", 
                                         min_value=1.0,
                                         value=default_h,
                                         step=0.1,
                                         format="%.1f",
                                         key=f"{key_prefix}_h",
                                         label_visibility="collapsed")
                
                # Save dimensions in session state
                st.session_state.file_dimensions[f"{filename}_L"] = L
                st.session_state.file_dimensions[f"{filename}_b"] = b
                st.session_state.file_dimensions[f"{filename}_h"] = h
                
                dimension_entries.append({
                    'filename': filename,
                    'L': L,
                    'b': b,
                    'h': h
                })
        
        st.divider()
        
        # Process each file
        for i, bend_file in enumerate(bend_files):
            filename = bend_file.name
            try:
                # Get dimensions for this file
                dims = next((d for d in dimension_entries if d['filename'] == filename), None)
                if not dims:
                    st.warning(f"No dimensions found for {filename}")
                    continue
                    
                L = dims['L']
                b = dims['b']
                h = dims['h']
                
                # Read CSV - handle trailing commas
                df = pd.read_csv(bend_file, header=None)
                df = df.dropna(axis=1, how='all')
                
                # Validate column count
                if len(df.columns) < 6:
                    st.error(f"File '{filename}' has only {len(df.columns)} columns. Expected at least 6 columns.")
                    continue
                    
                # Rename columns - use 6th column for force (index 5)
                df.columns = [f'col_{i}' for i in range(len(df.columns))]
                df['force_n'] = df.iloc[:, 5]
                
                # Clean data
                df = df.replace([np.inf, -np.inf], np.nan).dropna(subset=['force_n'])
                df['force_n'] = pd.to_numeric(df['force_n'], errors='coerce')
                df = df.dropna(subset=['force_n'])
                df = df[df['force_n'] >= 0]
                
                # Enhanced filtering
                mean_force = df['force_n'].mean()
                std_force = df['force_n'].std()
                if std_force > 0:
                    df = df[df['force_n'] <= mean_force + 3 * std_force]
                
                if not df.empty:
                    max_force = df['force_n'].max()
                    threshold = max_force * 0.01
                    start_index = df[df['force_n'] > threshold].index.min()
                    if pd.notnull(start_index):
                        df = df.loc[start_index:]
                
                if df.empty:
                    st.warning(f"File '{filename}' does not contain valid force data after cleaning")
                    continue
                
                # Find peak force
                max_force_n = df['force_n'].max()
                
                # Calculate bending strength
                L_cm = L * 0.1
                b_cm = b * 0.1
                h_cm = h * 0.1
                bending_strength = (3 * max_force_n * L_cm) / (2 * b_cm * h_cm**2)
                status = "✅ Pass" if bending_strength >= nominal_strength else "❌ Fail"
                
                # Extract part ID and job number from filename
                try:
                    date_part = filename.split("_")[0] + "_" + filename.split("_")[1]
                    identifier = filename.split("_")[2].split("(")[0]
                    job_no = filename.split("(")[1].split(")")[0] if "(" in filename else "N/A"
                    part_id = f"{date_part}_{identifier}"
                except:
                    part_id = "Unknown"
                    job_no = "N/A"

                # Store results
                results.append({
                    'Filename': filename,
                    'Part ID': part_id,
                    'Job No': job_no,
                    'L (mm)': L,
                    'b (mm)': b,
                    'h (mm)': h,
                    'Max Force (N)': max_force_n,
                    'Bending Strength (N/cm²)': bending_strength,
                    'Status': status
                })
                dfs.append((filename, df))

                # Display individual results
                st.subheader(f"Results for {filename}")
                col1, col2, col3 = st.columns(3)
                col1.metric("Dimensions", f"{L}×{b}×{h} mm")
                col2.metric("Maximum Force", f"{max_force_n:.2f} N")
                col3.metric("Bending Strength", f"{bending_strength:.2f} N/cm²", 
                            delta="Pass" if status == "✅ Pass" else "Fail",
                            delta_color="normal" if status == "✅ Pass" else "inverse")
                
                # Quality status with styling
                status_html = f'<div class="{"pass-metric" if status == "✅ Pass" else "fail-metric"}">{status}</div>'
                st.markdown(status_html, unsafe_allow_html=True)
                
                # Create force progression plot
                fig, ax = plt.subplots(figsize=(10, 4))
                ax.plot(df.index, df['force_n'], label='Force (N)', color='#00509d')
                ax.axhline(y=max_force_n, color='r', linestyle='--', label='Max Force')
                ax.set_xlabel('Data Point Index')
                ax.set_ylabel('Force (N)')
                ax.set_title(f'Force Progression for {filename}')
                ax.legend()
                ax.grid(True, linestyle='--', alpha=0.7)
                st.pyplot(fig)
                plt.close(fig)

                if bending_strength < nominal_strength:
                    st.warning("""
                    **Recommendations to Increase Strength:**
                    - Place parts in oven at 140°C for 3 hours
                    - Increase binder amount
                    - Extend rest period before testing
                    - Check for printing defects (layer separation)
                    """)

            except Exception as e:
                st.error(f"Error processing file {filename}: {str(e)}")
                st.info("""
                **Required CSV Format:**
                - Must have at least 6 columns
                - 6th column should contain force values in Newtons (N)
                - Example row: `-10.7649,0,0,1.064,1.064,10.7649`
                """)

        # Display summary table
        if results:
            st.subheader("Test Summary")
            summary_df = pd.DataFrame(results)
            st.dataframe(summary_df.style.format({
                'L (mm)': '{:.1f}',
                'b (mm)': '{:.1f}',
                'h (mm)': '{:.1f}',
                'Max Force (N)': '{:.2f}',
                'Bending Strength (N/cm²)': '{:.2f}'
            }).applymap(lambda x: 'background-color: #d4edda' if x == '✅ Pass' else 'background-color: #ffcccc', 
                      subset=['Status']))

            # Generate combined Excel report
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Front Page
                front_sheet = workbook.add_worksheet('Test Summary')
                title_format = workbook.add_format({
                    'font_name': 'Calibri', 'font_size': 18, 'bold': True,
                    'align': 'center', 'valign': 'vcenter', 'bottom': 6
                })
                header_format = workbook.add_format({
                    'font_name': 'Calibri', 'font_size': 12, 'bold': True,
                    'bg_color': '#003366', 'font_color': 'white',
                    'border': 1, 'align': 'center', 'valign': 'vcenter'
                })
                subheader_format = workbook.add_format({
                    'font_name': 'Calibri', 'font_size': 14, 'bold': True, 'align': 'left'
                })
                info_format = workbook.add_format({
                    'font_name': 'Calibri', 'font_size': 11, 'align': 'left', 'text_wrap': True
                })
                parameter_format = workbook.add_format({
                    'font_name': 'Calibri', 'font_size': 11, 'bold': True,
                    'align': 'left', 'bg_color': '#e6f0f9'
                })
                value_format = workbook.add_format({
                    'font_name': 'Calibri', 'font_size': 11, 'align': 'left', 'border': 1
                })
                number_format = workbook.add_format({
                    'font_name': 'Calibri', 'font_size': 11, 'num_format': '0.00',
                    'align': 'left', 'border': 1
                })
                
                front_sheet.set_column('A:A', 2)
                front_sheet.set_column('B:B', 25)
                front_sheet.set_column('C:C', 25)
                front_sheet.set_column('D:D', 2)
                
                front_sheet.merge_range('B1:D1', 'Brafe Engineering - Bend Test Report', title_format)
                front_sheet.merge_range('B3:D3', 'Quality Control Department', info_format)
                front_sheet.merge_range('B4:D4', f"Report Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", info_format)
                front_sheet.merge_range('B6:D6', 'Test Parameters', subheader_format)
                
                operator_name = st.session_state.get("operator_name", "Unknown Operator")
                test_id = st.session_state.get("test_id", "Unknown Test ID")
                test_date = datetime.datetime.now().strftime('%Y-%m-%d')
                
                parameters = [
                    ("Test ID", test_id),
                    ("Operator", operator_name),
                    ("Test Date", test_date),
                    ("Nominal Strength", f"{nominal_strength} N/cm²")
                ]
                
                for row_idx, (param, value) in enumerate(parameters, start=7):
                    front_sheet.write(row_idx, 1, param, parameter_format)
                    front_sheet.write(row_idx, 2, value, value_format)
                
                front_sheet.merge_range('B13:D13', 'Test Results Summary', subheader_format)
                
                if results:
                    headers = ["Filename", "Part ID", "Job No", "Strength (N/cm²)", "Status"]
                    for col_idx, header in enumerate(headers, start=1):
                        front_sheet.write(14, col_idx, header, header_format)
                    
                    for row_idx, result in enumerate(results, start=15):
                        front_sheet.write(row_idx, 1, result['Filename'], value_format)
                        front_sheet.write(row_idx, 2, result['Part ID'], value_format)
                        front_sheet.write(row_idx, 3, result['Job No'], value_format)
                        front_sheet.write_number(row_idx, 4, result['Bending Strength (N/cm²)'], number_format)
                        
                        if result['Status'] == "✅ Pass":
                            status_format = workbook.add_format({
                                'font_color': '#155724', 'bg_color': '#d4edda', 'border': 1
                            })
                        else:
                            status_format = workbook.add_format({
                                'font_color': '#721c24', 'bg_color': '#f8d7da', 'border': 1
                            })
                        front_sheet.write(row_idx, 5, result['Status'], status_format)
                    
                    strengths = [r['Bending Strength (N/cm²)'] for r in results]
                    if strengths:
                        avg_strength = sum(strengths) / len(strengths)
                        min_strength = min(strengths)
                        max_strength = max(strengths)
                        
                        front_sheet.merge_range(f'B{16+len(results)}:C{16+len(results)}', 'Average Strength', parameter_format)
                        front_sheet.write_number(15+len(results), 4, avg_strength, number_format)
                        
                        front_sheet.merge_range(f'B{17+len(results)}:C{17+len(results)}', 'Minimum Strength', parameter_format)
                        front_sheet.write_number(16+len(results), 4, min_strength, number_format)
                        
                        front_sheet.merge_range(f'B{18+len(results)}:C{18+len(results)}', 'Maximum Strength', parameter_format)
                        front_sheet.write_number(17+len(results), 4, max_strength, number_format)
                
                front_sheet.merge_range(f'B{20+len(results)}:D{20+len(results)}', "Note: Complete test data available in subsequent sheets", info_format)
                
                # Breaking Force Sheet
                if results:
                    breaking_force_sheet = workbook.add_worksheet('Breaking Force')
                    breaking_force_sheet.merge_range('A1:E1', 'Breaking Force Measurements', title_format)
                    breaking_force_sheet.write_row(2, 0, ['Filename', 'Part ID', 'Job No', 'Max Force (N)', 'Status'], header_format)
                    
                    for row_idx, result in enumerate(results, start=3):
                        breaking_force_sheet.write(row_idx, 0, result['Filename'], value_format)
                        breaking_force_sheet.write(row_idx, 1, result['Part ID'], value_format)
                        breaking_force_sheet.write(row_idx, 2, result['Job No'], value_format)
                        breaking_force_sheet.write_number(row_idx, 3, result['Max Force (N)'], number_format)
                        
                        if result['Status'] == "✅ Pass":
                            status_format = workbook.add_format({
                                'font_color': '#155724', 'bg_color': '#d4edda'
                            })
                        else:
                            status_format = workbook.add_format({
                                'font_color': '#721c24', 'bg_color': '#f8d7da'
                            })
                        breaking_force_sheet.write(row_idx, 4, result['Status'], status_format)
                    
                    breaking_force_sheet.set_column('A:A', 30)
                    breaking_force_sheet.set_column('B:B', 20)
                    breaking_force_sheet.set_column('C:C', 15)
                    breaking_force_sheet.set_column('D:D', 15)
                    breaking_force_sheet.set_column('E:E', 10)
                
                # Sheets for Each Test
                if results and dfs:
                    for idx, (filename, df) in enumerate(dfs):
                        base_name = re.sub(r'[\[\]:*?/\\]', '_', filename)[:20]
                        
                        # Parameter sheet
                        param_sheet_name = base_name + "_Params"
                        param_sheet = workbook.add_worksheet(param_sheet_name[:31])
                        param_sheet.merge_range('A1:C1', f'Test Parameters: {filename}', title_format)
                        param_sheet.write_row(2, 0, ['Parameter', 'Value', 'Status'], header_format)
                        
                        file_result = next((r for r in results if r['Filename'] == filename), {})
                        
                        test_params = [
                            ("Test ID", test_id, ""),
                            ("Operator", operator_name, ""),
                            ("Test Date", test_date, ""),
                            ("Part ID", file_result.get('Part ID', 'N/A'), ""),
                            ("Job No", file_result.get('Job No', 'N/A'), ""),
                            ("Support Span (L)", f"{file_result.get('L (mm)', 0):.1f} mm", ""),
                            ("Width (b)", f"{file_result.get('b (mm)', 0):.1f} mm", ""),
                            ("Height (h)", f"{file_result.get('h (mm)', 0):.1f} mm", ""),
                            ("Max Force", f"{file_result.get('Max Force (N)', 0):.2f} N", ""),
                            ("Bending Strength", f"{file_result.get('Bending Strength (N/cm²)', 0):.2f} N/cm²", ""),
                            ("Status", "", file_result.get('Status', 'N/A'))
                        ]
                        
                        for row_idx, (param, value, status) in enumerate(test_params, start=3):
                            param_sheet.write(row_idx, 0, param, parameter_format)
                            param_sheet.write(row_idx, 1, value, value_format)
                            
                            if status == "✅ Pass":
                                status_format = workbook.add_format({
                                    'font_color': '#155724', 'bg_color': '#d4edda', 'border': 1
                                })
                            else:
                                status_format = workbook.add_format({
                                    'font_color': '#721c24', 'bg_color': '#f8d7da', 'border': 1
                                })
                            param_sheet.write(row_idx, 2, status, status_format)
                        
                        param_sheet.set_column('A:A', 25)
                        param_sheet.set_column('B:B', 20)
                        param_sheet.set_column('C:C', 15)
                        
                        # Raw data sheet
                        data_sheet_name = base_name + "_Data"
                        raw_sheet = workbook.add_worksheet(data_sheet_name[:31])
                        raw_sheet.merge_range('A1:Z1', f'Raw Test Data: {filename}', title_format)
                        headers = df.columns.tolist()
                        raw_sheet.write_row(2, 0, headers, header_format)
                        
                        for row_idx, row in enumerate(df.values):
                            for col_idx, value in enumerate(row):
                                if isinstance(value, float):
                                    raw_sheet.write_number(row_idx + 3, col_idx, value, number_format)
                                else:
                                    raw_sheet.write_string(row_idx + 3, col_idx, str(value), value_format)
                        
                        raw_sheet.freeze_panes(3, 0)
                        
                        if not df.empty:
                            chart = workbook.add_chart({'type': 'line'})
                            chart.add_series({
                                'values': [data_sheet_name, 3, 5, 3 + len(df), 5],
                                'name': 'Force (N)',
                                'line': {'color': '#003366', 'width': 1.5}
                            })
                            
                            max_force = df['force_n'].max()
                            max_index = df['force_n'].idxmax() + 3
                            
                            chart.add_series({
                                'values': [data_sheet_name, max_index, 5, max_index, 5],
                                'name': 'Max Force',
                                'marker': {'type': 'circle', 'size': 6, 'fill': {'color': '#FF0000'}},
                                'line': {'none': True}
                            })
                            
                            chart.set_title({'name': f'Force Progression: {filename}'})
                            chart.set_x_axis({'name': 'Data Point Index'})
                            chart.set_y_axis({'name': 'Force (N)'})
                            chart.set_legend({'position': 'top'})
                            
                            raw_sheet.insert_chart(f'G{len(df) + 10}', chart)

            st.download_button(
                label="📥 Download Excel Report",
                data=output.getvalue(),
                file_name=f"Brafe_BendTest_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download comprehensive test report in Excel format"
            )

with tab3:
    st.header("Loss on Ignition (LOI) Analysis")
    st.caption("Calculate binder content according to section 3.5 of Quality Control Manual")
    
    method = st.radio("Test Method", 
                      ["Bunsen Burner (Section 3.5.1)", "Oven (Section 3.5.2)"],
                      index=0,
                      horizontal=True)
    
    with st.form("loi_calculation"):
        st.subheader("Enter Measurement Values")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            t1 = st.number_input("T1 (Bowl Weight) in g", 
                                 value=-44.904,
                                 format="%.3f",
                                 help="Negative value shown on scale after taring")
        with col2:
            w1 = st.number_input("W1 (Sample Weight) in g", 
                                 min_value=20.0,
                                 max_value=40.0,
                                 value=30.023,
                                 format="%.3f")
        with col3:
            t2 = st.number_input("T2 (Bowl + Ash) in g", 
                                 value=-74.422,
                                 format="%.3f",
                                 help="Negative value shown on scale after taring")
        
        submitted = st.form_submit_button("Calculate LOI")
        
        if submitted:
            try:
                delta_m = (abs(t2) - abs(t1)) - w1
                loi = (abs(delta_m) / w1) * 100
                
                st.session_state.loi_results = {
                    'delta_m': abs(delta_m),
                    'loi': loi,
                    'status': "✅ Pass" if 0.5 <= loi <= 2.5 else "❌ Fail",
                    't1': t1,
                    'w1': w1,
                    't2': t2
                }
                
                st.divider()
                st.subheader("Results")
                
                col1, col2 = st.columns(2)
                col1.metric("Mass Loss (Δm)", f"{abs(delta_m):.3f} g")
                col2.metric("Loss on Ignition", f"{loi:.2f} %")
                
                status = "✅ Pass" if 0.5 <= loi <= 2.5 else "❌ Fail"
                st.markdown(f'<div class="{"pass-metric" if status == "✅ Pass" else "fail-metric"}">{status} - {"Optimal binder content" if status == "✅ Pass" else "Out of optimal range"}</div>', 
                            unsafe_allow_html=True)
                
                st.info("""
                **Interpretation Guide:**
                - Optimal range: 0.5-2.5%
                - < 0.5%: Insufficient binder
                - > 2.5%: Excessive binder
                """)
                
                if "Bunsen" in method:
                    st.caption("Bunsen Burner Method Notes:\n- Burn until sand turns white\n- Stir every minute\n- Cool for 20 min before weighing")
                else:
                    st.caption("Oven Method Notes:\n- Heat to 900°C for 3 hours\n- Cool in closed oven before weighing")
                    
            except Exception as e:
                st.error(f"Calculation error: {str(e)}")
    
    # Download button outside the form
    if st.session_state.loi_results:
        operator_name = st.session_state.get("operator_name", "Unknown Operator")
        test_id = st.session_state.get("test_id", "Unknown Test ID")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            summary_df = pd.DataFrame({
                'Parameter': ['Test Date', 'Operator', 'Test ID', 
                              'Method', 'T1 (g)', 'W1 (g)', 'T2 (g)',
                              'Mass Loss (g)', 'LOI (%)', 'Status'],
                'Value': [datetime.datetime.now().strftime('%Y-%m-%d'), 
                          operator_name, 
                          test_id,
                          method, 
                          st.session_state.loi_results['t1'], 
                          st.session_state.loi_results['w1'], 
                          st.session_state.loi_results['t2'],
                          st.session_state.loi_results['delta_m'], 
                          st.session_state.loi_results['loi'], 
                          st.session_state.loi_results['status']]
            })
            summary_df.to_excel(writer, sheet_name='Test Summary', index=False, startrow=1)
            
            workbook = writer.book
            summary_sheet = writer.sheets['Test Summary']
            
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#003366',
                'font_color': 'white',
                'border': 1
            })
            pass_format = workbook.add_format({
                'bg_color': '#d4edda',
                'font_color': '#155724'
            })
            fail_format = workbook.add_format({
                'bg_color': '#f8d7da',
                'font_color': '#721c24'
            })
            
            for col_num, value in enumerate(summary_df.columns.values):
                summary_sheet.write(0, col_num, value, header_format)
            
            status_row = summary_df.index[summary_df['Parameter'] == 'Status'].tolist()[0] + 1
            if st.session_state.loi_results['status'] == "✅ Pass":
                summary_sheet.conditional_format(f'B{status_row+1}', {
                    'type': 'cell',
                    'criteria': '==',
                    'value': '"✅ Pass"',
                    'format': pass_format
                })
            else:
                summary_sheet.conditional_format(f'B{status_row+1}', {
                    'type': 'cell',
                    'criteria': '==',
                    'value': '"❌ Fail"',
                    'format': fail_format
                })
            
            summary_sheet.set_column('A:A', 25)
            summary_sheet.set_column('B:B', 20)
        
        st.download_button(
            label="📥 Download Excel Report",
            data=output.getvalue(),
            file_name=f"Brafe_LOI_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download LOI test report in Excel format"
        )
    
    st.divider()
    st.subheader("LOI Formula Reference")
    st.latex(r'''
    \begin{align*}
    \Delta m &= (|T2| - |T1|) - W1 \\
    LOI (\%) &= \left( \frac{|\Delta m|}{W1} \right) \times 100
    \end{align*}
    ''')
    st.caption("Note: Algebraic signs are not considered in calculations (per manual section 3.5)")

# Footer with Brafe branding
st.divider()
st.caption("""
**Quality Control Manual Reference:** PDB_02P06PDBQL2 (Version 0001, Dec 2022) 
""")
st.caption("© 2023 Brafe Engineering | All Rights Reserved")
