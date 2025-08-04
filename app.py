import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import base64
from io import BytesIO
import datetime
import re
import os

# Load measurement images and Brafe logo with error handling
try:
    x_img = Image.open('x_measurement.png')
except FileNotFoundError:
    x_img = None
    st.warning("Image 'x_measurement.png' not found. Using placeholder.")
try:
    y_img = Image.open('y_measurement.png')
except FileNotFoundError:
    y_img = None
    st.warning("Image 'y_measurement.png' not found. Using placeholder.")
try:
    z_img = Image.open('z_measurement.png')
except FileNotFoundError:
    z_img = None
    st.warning("Image 'z_measurement.png' not found. Using placeholder.")
try:
    brafe_logo = Image.open('brafe_logo.png')
except FileNotFoundError:
    brafe_logo = None
    st.warning("Image 'brafe_logo.png' not found. Using placeholder.")

# App configuration with updated blue theme
st.set_page_config(
    page_title="Brafe Engineering Quality Control",
    page_icon=brafe_logo if brafe_logo else ":bar_chart:",
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

# Sidebar with Brafe branding
with st.sidebar:
    # Brafe logo
    if brafe_logo:
        st.image(brafe_logo, width=150)
    else:
        st.markdown("**Brafe Logo Placeholder**")
    
    # Test Information
    st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
    st.subheader("Test Information")
    st.session_state.operator_name = st.text_input("Operator Name", "John Doe", key="global_operator")
    st.session_state.test_id = st.text_input("Test ID", "TEST-001", key="global_test_id")
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Engineering Resources
    st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
    st.header("Brafe Engineering Resources")
    st.markdown("[Quality Control Manual](https://www.brafeengineering.com/support/)")
    st.markdown("[Technical Support](mailto:support@brafeengineering.com)")
    st.markdown("Hotline: +44 123 456 7890")
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
            if x_img:
                st.image(x_img, caption="Measure Dimension X (Length)", use_container_width=True)
            else:
                st.markdown("**Image Placeholder: Measure Dimension X (Length)**")
            st.info("**X-Dimension:**\n- Length direction\n- Nominal: 172 mm")
        with cols[1]:
            if y_img:
                st.image(y_img, caption="Measure Dimension Y (Width)", use_container_width=True)
            else:
                st.markdown("**Image Placeholder: Measure Dimension Y (Width)**")
            st.info("**Y-Dimension:**\n- Width direction\n- Nominal: 22.4 mm")
        with cols[2]:
            if z_img:
                st.image(z_img, caption="Measure Dimension Z (Height)", use_container_width=True)
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
    
    # Initialize session state for dimensions
    if 'file_dimensions' not in st.session_state:
        st.session_state.file_dimensions = {}
    
    bend_files = st.file_uploader("Upload CSV files from bend test machine",
                                  type=["csv"],
                                  accept_multiple_files=True,
                                  help="Should contain force measurements in last column (Newtons)")
    
    if bend_files:
        # Initialize lists to store results
        results = []
        dfs = []
        dimension_entries = []
        graphs = []  # Store graph images for Excel
        
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
                
                # Fix for files with trailing commas (like the example)
                # Remove any empty columns at the end
                df = df.dropna(axis=1, how='all')
                
                # Validate column count
                if len(df.columns) < 6:
                    st.error(f"File '{filename}' has only {len(df.columns)} columns. Expected at least 6 columns.")
                    continue
                    
                # Rename columns - use 6th column for force (index 5)
                df.columns = [f'col_{i}' for i in range(len(df.columns))]
                df['force_n'] = df.iloc[:, 5]  # Use 6th column as force
                
                # Clean data
                df = df.replace([np.inf, -np.inf], np.nan).dropna(subset=['force_n'])
                df['force_n'] = pd.to_numeric(df['force_n'], errors='coerce')
                df = df.dropna(subset=['force_n'])
                
                # Enhanced filtering
                # 1. Remove negative values (force should be positive)
                df = df[df['force_n'] >= 0]
                
                # 2. Remove extreme outliers (more than 3 std devs from mean)
                mean_force = df['force_n'].mean()
                std_force = df['force_n'].std()
                if std_force > 0:  # Avoid division by zero
                    df = df[df['force_n'] <= mean_force + 3 * std_force]
                
                # 3. Remove values that are too high before the main test starts
                # Find the first significant force value (>1% of max)
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
                
                # Calculate bending strength in N/cm²
                # Formula: σ = (3 * F * L) / (2 * b * h²)
                # Convert mm to cm: 1 mm = 0.1 cm
                L_cm = L * 0.1
                b_cm = b * 0.1
                h_cm = h * 0.1
                
                bending_strength = (3 * max_force_n * L_cm) / (2 * b_cm * h_cm**2)
                status = "✅ Pass" if bending_strength >= nominal_strength else "❌ Fail"
                
                # Extract part ID and job number from filename
                try:
                    # Example filename: 2025_0731_1110221A(1).csv
                    # Extract date and identifier
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
                
                # Save plot for Excel
                img_buffer = BytesIO()
                plt.savefig(img_buffer, format='png', dpi=300)
                plt.close(fig)
                graphs.append((filename, img_buffer))

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
                # Get the workbook object
                workbook = writer.book
                
                # ===================================================
                # ========== Create Professional Front Page ==========
                # ===================================================
                front_sheet = workbook.add_worksheet('Test Summary')
                
                # Define professional formats
                title_format = workbook.add_format({
                    'font_name': 'Calibri',
                    'font_size': 18,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'bottom': 6
                })
                
                header_format = workbook.add_format({
                    'font_name': 'Calibri',
                    'font_size': 12,
                    'bold': True,
                    'bg_color': '#003366',
                    'font_color': 'white',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter'
                })
                
                subheader_format = workbook.add_format({
                    'font_name': 'Calibri',
                    'font_size': 14,
                    'bold': True,
                    'align': 'left'
                })
                
                info_format = workbook.add_format({
                    'font_name': 'Calibri',
                    'font_size': 11,
                    'align': 'left',
                    'text_wrap': True
                })
                
                parameter_format = workbook.add_format({
                    'font_name': 'Calibri',
                    'font_size': 11,
                    'bold': True,
                    'align': 'left',
                    'bg_color': '#e6f0f9'
                })
                
                value_format = workbook.add_format({
                    'font_name': 'Calibri',
                    'font_size': 11,
                    'align': 'left',
                    'border': 1
                })
                
                number_format = workbook.add_format({
                    'font_name': 'Calibri',
                    'font_size': 11,
                    'num_format': '0.00',
                    'align': 'left',
                    'border': 1
                })
                
                # Set column widths
                front_sheet.set_column('A:A', 2)  # Padding
                front_sheet.set_column('B:B', 25)  # Labels
                front_sheet.set_column('C:C', 25)  # Values
                front_sheet.set_column('D:D', 2)  # Padding
                
                # Add title and company info
                front_sheet.merge_range('B1:D1', 'Brafe Engineering - Bend Test Report', title_format)
                front_sheet.merge_range('B3:D3', 'Quality Control Department', info_format)
                front_sheet.merge_range('B4:D4', f"Report Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", info_format)
                
                # Add test parameters section
                front_sheet.merge_range('B6:D6', 'Test Parameters', subheader_format)
                
                # Get operator and test ID from session state
                operator_name = st.session_state.get("operator_name", "Unknown Operator")
                test_id = st.session_state.get("test_id", "Unknown Test ID")
                
                # Add parameter table
                test_date = datetime.datetime.now().strftime('%Y-%m-%d')
                parameters = [
                    ("Test ID", test_id),
                    ("Operator", operator_name),
                    ("Test Date", test_date),
                    ("Nominal Strength", f"{nominal_strength} N/cm²"),
                    ("Tolerance", "±0.45 mm")
                ]
                
                for row_idx, (param, value) in enumerate(parameters, start=7):
                    front_sheet.write(row_idx, 1, param, parameter_format)
                    front_sheet.write(row_idx, 2, value, value_format)
                
                # Add test summary section
                front_sheet.merge_range('B13:D13', 'Test Results Summary', subheader_format)
                
                # Add summary table if results exist
                if results:
                    # Write headers
                    headers = ["Filename", "Part ID", "Job No", "Strength (N/cm²)", "Status"]
                    for col_idx, header in enumerate(headers, start=1):
                        front_sheet.write(14, col_idx, header, header_format)
                    
                    # Write data
                    for row_idx, result in enumerate(results, start=15):
                        front_sheet.write(row_idx, 1, result['Filename'], value_format)
                        front_sheet.write(row_idx, 2, result['Part ID'], value_format)
                        front_sheet.write(row_idx, 3, result['Job No'], value_format)
                        front_sheet.write_number(row_idx, 4, result['Bending Strength (N/cm²)'], number_format)
                        
                        # Apply conditional formatting for status
                        if result['Status'] == "✅ Pass":
                            status_format = workbook.add_format({
                                'font_color': '#155724',
                                'bg_color': '#d4edda',
                                'border': 1
                            })
                        else:
                            status_format = workbook.add_format({
                                'font_color': '#721c24',
                                'bg_color': '#f8d7da',
                                'border': 1
                            })
                        front_sheet.write(row_idx, 5, result['Status'], status_format)
                    
                    # Add statistics
                    strengths = [r['Bending Strength (N/cm²)'] for r in results]
                    if strengths:
                        avg_strength = sum(strengths) / len(strengths)
                        min_strength = min(strengths)
                        max_strength = max(strengths)
                        
                        front_sheet.merge_range(f'B{16+len(results)}:C{16+len(results)}', 'Average Strength', parameter_format)
                        front_sheet.write_number(16+len(results), 4, avg_strength, number_format)
                        
                        front_sheet.merge_range(f'B{17+len(results)}:C{17+len(results)}', 'Minimum Strength', parameter_format)
                        front_sheet.write_number(17+len(results), 4, min_strength, number_format)
                        
                        front_sheet.merge_range(f'B{18+len(results)}:C{18+len(results)}', 'Maximum Strength', parameter_format)
                        front_sheet.write_number(18+len(results), 4, max_strength, number_format)
                
                # Add footer note
                note = "Note: Complete test data available in subsequent sheets"
                front_sheet.merge_range(f'B{20+len(results)}:D{20+len(results)}', note, info_format)
                
                # ===================================================
                # ========== Create Breaking Force Sheet =============
                # ===================================================
                if results:
                    breaking_force_sheet = workbook.add_worksheet('Breaking Force')
                    
                    # Create header
                    breaking_force_sheet.merge_range('A1:E1', 'Breaking Force Measurements', title_format)
                    breaking_force_sheet.write_row(2, 0, ['Filename', 'Part ID', 'Job No', 'Max Force (N)', 'Status'], header_format)
                    
                    # Write data
                    for row_idx, result in enumerate(results, start=3):
                        breaking_force_sheet.write(row_idx, 0, result['Filename'], value_format)
                        breaking_force_sheet.write(row_idx, 1, result['Part ID'], value_format)
                        breaking_force_sheet.write(row_idx, 2, result['Job No'], value_format)
                        breaking_force_sheet.write_number(row_idx, 3, result['Max Force (N)'], number_format)
                        
                        # Apply conditional formatting for status
                        if result['Status'] == "✅ Pass":
                            status_format = workbook.add_format({
                                'font_color': '#155724',
                                'bg_color': '#d4edda'
                            })
                        else:
                            status_format = workbook.add_format({
                                'font_color': '#721c24',
                                'bg_color': '#f8d7da'
                            })
                        breaking_force_sheet.write(row_idx, 4, result['Status'], status_format)
                    
                    # Format columns
                    breaking_force_sheet.set_column('A:A', 30)
                    breaking_force_sheet.set_column('B:B', 20)
                    breaking_force_sheet.set_column('C:C', 15)
                    breaking_force_sheet.set_column('D:D', 15)
                    breaking_force_sheet.set_column('E:E', 10)
                
                # ===================================================
                # ========== Create Sheets for Each Test ============
                # ===================================================
                if results and dfs:
                    for idx, (filename, df) in enumerate(dfs):
                        # Create base name for sheets
                        base_name = re.sub(r'[\[\]:*?/\\]', '_', filename)[:20]
                        
                        # Create parameter sheet
                        param_sheet_name = base_name + "_Params"
                        param_sheet = workbook.add_worksheet(param_sheet_name[:31])
                        
                        # Add header
                        param_sheet.merge_range('A1:C1', f'Test Parameters: {filename}', title_format)
                        param_sheet.write_row(2, 0, ['Parameter', 'Value', 'Status'], header_format)
                        
                        # Get specific values for this file
                        file_result = next((r for r in results if r['Filename'] == filename), {})
                        
                        # Prepare parameters
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
                        
                        # Write parameters
                        for row_idx, (param, value, status) in enumerate(test_params, start=3):
                            param_sheet.write(row_idx, 0, param, parameter_format)
                            param_sheet.write(row_idx, 1, value, value_format)
                            
                            # Apply status formatting
                            if status == "✅ Pass":
                                status_format = workbook.add_format({
                                    'font_color': '#155724',
                                    'bg_color': '#d4edda',
                                    'border': 1
                                })
                            else:
                                status_format = workbook.add_format({
                                    'font_color': '#721c24',
                                    'bg_color': '#f8d7da',
                                    'border': 1
                                })
                            param_sheet.write(row_idx, 2, status, status_format)
                        
                        # Set column widths
                        param_sheet.set_column('A:A', 25)
                        param_sheet.set_column('B:B', 20)
                        param_sheet.set_column('C:C', 15)
                        
                        # Create raw data sheet
                        data_sheet_name = base_name + "_Data"
                        raw_sheet = workbook.add_worksheet(data_sheet_name[:31])
                        
                        # Add header
                        raw_sheet.merge_range('A1:Z1', f'Raw Test Data: {filename}', title_format)
                        
                        # Write raw data headers
                        headers = df.columns.tolist()
                        raw_sheet.write_row(2, 0, headers, header_format)
                        
                        # Write raw data
                        for row_idx, row in enumerate(df.values):
                            for col_idx, value in enumerate(row):
                                if isinstance(value, float):
                                    raw_sheet.write_number(row_idx + 3, col_idx, value, number_format)
                                else:
                                    raw_sheet.write_string(row_idx + 3, col_idx, str(value), value_format)
                        
                        # Freeze header row
                        raw_sheet.freeze_panes(3, 0)
                        
                        # Add chart
                        if not df.empty:
                            chart = workbook.add_chart({'type': 'line'})
                            chart.add_series({
                                'values': [data_sheet_name, 3, 5, 3 + len(df), 5],
                                'name': 'Force (N)',
                                'line': {'color': '#003366', 'width': 1.5}
                            })
                            
                            # Find max force value and position
                            max_force = df['force_n'].max()
                            max_index = df['force_n'].idxmax() + 3  # +3 for header offset
                            
                            # Add max force marker
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
                            
                            # Insert chart below data
                            raw_sheet.insert_chart(f'G{len(df) + 10}', chart)

            # Only show download button if we have results
            if results:
                st.download_button(
                    label="📥 Download Excel Report",
                    data=output.getvalue(),
                    file_name=f"Brafe_BendTest_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download comprehensive test report in Excel format"
                )
            else:
                st.warning("No test results available to generate report")
         
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
                # The manual's formula, which uses absolute values of tared bowl weights,
                # is not standard but is implemented as specified.
                delta_m = (abs(t2) - abs(t1)) - w1
                loi = (abs(delta_m) / w1) * 100
                
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
                
                # Get operator and test ID from session state
                operator_name = st.session_state.get("operator_name", "Unknown Operator")
                test_id = st.session_state.get("test_id", "Unknown Test ID")
                
                # Generate Excel report in Brafe template format
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Create summary sheet
                    summary_df = pd.DataFrame({
                        'Parameter': ['Test Date', 'Operator', 'Test ID', 
                                      'Method', 'T1 (g)', 'W1 (g)', 'T2 (g)',
                                      'Mass Loss (g)', 'LOI (%)', 'Status'],
                        'Value': [datetime.datetime.now().strftime('%Y-%m-%d'), 
                                  operator_name, 
                                  test_id,
                                  method, t1, w1, t2,
                                  abs(delta_m), loi, status]
                    })
                    summary_df.to_excel(writer, sheet_name='Test Summary', index=False, startrow=1)
                    
                    # Formatting
                    workbook = writer.book
                    summary_sheet = writer.sheets['Test Summary']
                    
                    # Formatting
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
                    
                    # Apply header formatting
                    for col_num, value in enumerate(summary_df.columns.values):
                        summary_sheet.write(0, col_num, value, header_format)
                    
                    # Apply conditional formatting to status
                    status_row = summary_df.index[summary_df['Parameter'] == 'Status'].tolist()[0] + 1
                    if status == "✅ Pass":
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
                    
                    # Set column widths
                    summary_sheet.set_column('A:A', 25)
                    summary_sheet.set_column('B:B', 20)
                
                # Download button
                st.download_button(
                    label="📥 Download Excel Report",
                    data=output.getvalue(),
                    file_name=f"Brafe_LOI_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download LOI test report in Excel format"
                )
                
                if "Bunsen" in method:
                    st.caption("Bunsen Burner Method Notes:\n- Burn until sand turns white\n- Stir every minute\n- Cool for 20 min before weighing")
                else:
                    st.caption("Oven Method Notes:\n- Heat to 900°C for 3 hours\n- Cool in closed oven before weighing")
                    
            except Exception as e:
                st.error(f"Calculation error: {str(e)}")
    
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
