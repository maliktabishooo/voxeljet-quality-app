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
    </style>
    """,
    unsafe_allow_html=True
)

# Sidebar with Brafe branding
with st.sidebar:
    if brafe_logo:
        st.image(brafe_logo, width=150)
    else:
        st.markdown("**Brafe Logo Placeholder**")
    st.header("Brafe Engineering Resources")
    st.markdown("[Quality Control Manual](https://www.brafeengineering.com/support/)")
    st.markdown("[Technical Support](mailto:support@brafeengineering.com)")
    st.markdown("Hotline: +44 123 456 7890")
    
    st.divider()
    st.caption("Test Specifications")
    st.metric("Nominal Bending Strength", "260 N/cm²")
    st.metric("Test Bar Dimensions", "172 × 22.4 × 22.4 mm")
    st.metric("LOI Sample Weight", "30 g")
    st.metric("Dimensional Tolerance", "±0.45 mm")

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
            
            # New code: Add total dimension metric
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

    # Manual Force Input Option
    manual_mode = st.checkbox("Enter Force Value Manually", value=False)
    
    # Default dimensions for reference
    DEFAULT_L = 172.0
    DEFAULT_B = 22.4
    DEFAULT_H = 22.4
    
    if manual_mode:
        # Manual mode with per-bar dimensions
        with st.form("manual_force"):
            st.subheader("Enter Test Bar Parameters")
            
            # Create columns for multiple bars
            num_bars = st.number_input("Number of Test Bars", min_value=1, max_value=10, value=1)
            
            manual_results = []
            cols = st.columns(3)
            
            with cols[0]:
                st.subheader("Support Span (L mm)")
            with cols[1]:
                st.subheader("Width (b mm)")
            with cols[2]:
                st.subheader("Height (h mm)")
            
            for i in range(num_bars):
                cols = st.columns(3)
                with cols[0]:
                    L = st.number_input(f"Bar {i+1} Span", 
                                      value=DEFAULT_L, 
                                      min_value=1.0, 
                                      key=f"L_{i}",
                                      format="%.1f")
                with cols[1]:
                    b = st.number_input(f"Bar {i+1} Width", 
                                      value=DEFAULT_B, 
                                      min_value=1.0, 
                                      key=f"b_{i}",
                                      format="%.1f")
                with cols[2]:
                    h = st.number_input(f"Bar {i+1} Height", 
                                      value=DEFAULT_H, 
                                      min_value=1.0, 
                                      key=f"h_{i}",
                                      format="%.1f")
            
            st.divider()
            st.subheader("Enter Force Values")
            
            force_cols = st.columns(num_bars)
            forces = []
            for i in range(num_bars):
                with force_cols[i]:
                    force = st.number_input(f"Force for Bar {i+1} (N)", 
                                          min_value=0.0, 
                                          value=285.76,
                                          key=f"force_{i}",
                                          format="%.2f")
                    forces.append(force)
            
            submitted = st.form_submit_button("Calculate Bending Strength")
            
            if submitted:
                results = []
                for i in range(num_bars):
                    # Calculate bending strength in N/cm²
                    L_cm = L * 0.1
                    b_cm = b * 0.1
                    h_cm = h * 0.1
                    
                    bending_strength = (3 * forces[i] * L_cm) / (2 * b_cm * h_cm**2)
                    status = "✅ Pass" if bending_strength >= 260 else "❌ Fail"
                    
                    results.append({
                        'Bar #': i+1,
                        'Support Span (mm)': L,
                        'Width (mm)': b,
                        'Height (mm)': h,
                        'Max Force (N)': forces[i],
                        'Bending Strength (N/cm²)': bending_strength,
                        'Status': status
                    })
                
                # Display results
                st.subheader("Results")
                result_df = pd.DataFrame(results)
                st.dataframe(result_df.style.format({
                    'Max Force (N)': '{:.2f}',
                    'Bending Strength (N/cm²)': '{:.2f}'
                }))
                
                # Show status metrics
                for i, res in enumerate(results):
                    cols = st.columns(5)
                    cols[0].metric("Bar", f"{i+1}")
                    cols[1].metric("Force", f"{res['Max Force (N)']:.2f} N")
                    cols[2].metric("Strength", f"{res['Bending Strength (N/cm²)']:.2f} N/cm²")
                    cols[3].metric("Status", res['Status'])
                    cols[4].metric("Target", "260 N/cm²")
                    
                    if res['Bending Strength (N/cm²)'] < 260:
                        st.warning(f"**Bar {i+1} Recommendations:** Increase binder amount or extend rest period")
                
                # Download button
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, sheet_name='Results', index=False)
                    
                    # Formatting
                    workbook = writer.book
                    worksheet = writer.sheets['Results']
                    
                    header_format = workbook.add_format({
                        'bold': True,
                        'bg_color': '#003366',
                        'font_color': 'white',
                        'border': 1
                    })
                    
                    for col_num, value in enumerate(result_df.columns):
                        worksheet.write(0, col_num, value, header_format)
                    
                    worksheet.set_column('A:A', 10)
                    worksheet.set_column('B:D', 20)
                    worksheet.set_column('E:F', 25)
                    worksheet.set_column('G:G', 15)
                
                st.download_button(
                    label="Download Excel Report",
                    data=output.getvalue(),
                    file_name=f"Brafe_Manual_BendTest_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        # CSV mode with per-bar dimensions
        st.subheader("Upload Bend Test Data")
        bend_files = st.file_uploader("Upload CSV files from bend test machine",
                                      type=["csv"],
                                      accept_multiple_files=True,
                                      help="Should contain force measurements in last column (Newtons)")

        if bend_files:
            # Initialize lists to store results for summary table
            results = []
            dfs = []
            
            # Section for per-file dimensions
            st.subheader("Specify Dimensions for Each Test Bar")
            st.caption("Enter custom dimensions for each test bar or use global defaults")
            
            # Global defaults
            with st.expander("⚙️ Global Default Dimensions", expanded=False):
                global_cols = st.columns(3)
                with global_cols[0]:
                    global_L = st.number_input("**Default Support Span (L)**", 
                                             value=DEFAULT_L, 
                                             min_value=1.0, 
                                             format="%.1f")
                with global_cols[1]:
                    global_b = st.number_input("**Default Width (b)**", 
                                             value=DEFAULT_B, 
                                             min_value=1.0, 
                                             format="%.1f")
                with global_cols[2]:
                    global_h = st.number_input("**Default Height (h)**", 
                                             value=DEFAULT_H, 
                                             min_value=1.0, 
                                             format="%.1f")
            
            # Create dimension inputs for each file
            file_params = {}
            for i, bend_file in enumerate(bend_files):
                filename = bend_file.name
                
                cols = st.columns(4)
                with cols[0]:
                    st.markdown(f"**{filename}**")
                with cols[1]:
                    L = st.number_input(f"Span for {filename}", 
                                      value=global_L, 
                                      min_value=1.0, 
                                      key=f"L_{filename}",
                                      format="%.1f")
                with cols[2]:
                    b = st.number_input(f"Width for {filename}", 
                                      value=global_b, 
                                      min_value=1.0, 
                                      key=f"b_{filename}",
                                      format="%.1f")
                with cols[3]:
                    h = st.number_input(f"Height for {filename}", 
                                      value=global_h, 
                                      min_value=1.0, 
                                      key=f"h_{filename}",
                                      format="%.1f")
                
                file_params[filename] = {'L': L, 'b': b, 'h': h}
            
            if st.button("Process All Files", key="process_all"):
                for bend_file in bend_files:
                    filename = bend_file.name
                    params = file_params[filename]
                    
                    try:
                        # Extract part ID and job number from filename
                        path_split_ext = os.path.splitext(filename)
                        path_split = os.path.split(path_split_ext[0])
                        test_key = path_split[1][16:] if len(path_split[1]) >= 16 else path_split[1]

                        try:
                            part_id = test_key.split("(")[0]
                            job_no = test_key.split('(')[1][:-1]
                        except:
                            part_id = "Unknown"
                            job_no = "N/A"

                        # Read CSV - using last column as force (Newtons)
                        df = pd.read_csv(bend_file, header=None)
                        
                        # Validate column count
                        if len(df.columns) < 6:
                            st.error(f"File '{filename}' has only {len(df.columns)} columns. Expected at least 6 columns.")
                            continue
                            
                        # Rename columns - focus on last column for force
                        df.columns = [f'col_{i}' for i in range(len(df.columns))]
                        df['force_n'] = df.iloc[:, -1]  # Use last column as force
                        
                        # Clean data - filter bad sensor readings
                        # Step 1: Remove invalid values
                        df = df.replace([np.inf, -np.inf], np.nan)
                        df = df.dropna(subset=['force_n'])
                        
                        # Step 2: Convert to numeric
                        df['force_n'] = pd.to_numeric(df['force_n'], errors='coerce')
                        df = df.dropna(subset=['force_n'])
                        
                        # Step 3: Filter negative values
                        df = df[df['force_n'] >= 0]
                        
                        # Step 4: Remove outliers using IQR method
                        Q1 = df['force_n'].quantile(0.25)
                        Q3 = df['force_n'].quantile(0.75)
                        IQR = Q3 - Q1
                        lower_bound = Q1 - 1.5 * IQR
                        upper_bound = Q3 + 1.5 * IQR
                        
                        filtered_df = df[(df['force_n'] >= lower_bound) & 
                                         (df['force_n'] <= upper_bound)]
                        
                        # If too much data removed, use original
                        if len(filtered_df) < 0.7 * len(df):
                            st.warning(f"Outlier filtering removed >30% of data for {filename}. Using unfiltered data.")
                            filtered_df = df.copy()
                        
                        # Step 5: Find start of test (first significant force value)
                        force_threshold = 0.1 * filtered_df['force_n'].max()
                        start_index = filtered_df[filtered_df['force_n'] > force_threshold].index[0] if not filtered_df[filtered_df['force_n'] > force_threshold].empty else 0
                        test_df = filtered_df.iloc[start_index:].reset_index(drop=True)
                        
                        # Step 6: Smooth force data
                        test_df['smoothed_force_n'] = test_df['force_n'].rolling(window=5, min_periods=1).mean()
                        max_force_n = test_df['smoothed_force_n'].max()
                        
                        # Calculate bending strength in N/cm²
                        # Formula: σ = (3 * F * L) / (2 * b * h²)
                        L_cm = params['L'] * 0.1
                        b_cm = params['b'] * 0.1
                        h_cm = params['h'] * 0.1
                        
                        bending_strength = (3 * max_force_n * L_cm) / (2 * b_cm * h_cm**2)
                        status = "✅ Pass" if bending_strength >= 260 else "❌ Fail"

                        # Store results
                        results.append({
                            'Filename': filename,
                            'Part ID': part_id,
                            'Job No': job_no,
                            'Support Span (mm)': params['L'],
                            'Width (mm)': params['b'],
                            'Height (mm)': params['h'],
                            'Max Force (N)': max_force_n,
                            'Bending Strength (N/cm²)': bending_strength,
                            'Status': status
                        })
                        dfs.append((filename, test_df))

                    except Exception as e:
                        st.error(f"Error processing file {filename}: {str(e)}")
                        st.info("""
                        **Required CSV Format:**
                        - Must have at least 6 columns
                        - Last column should contain force values in Newtons (N)
                        - Example row: `-10.7649,0,0,1.064,1.064,10.7649`
                        """)

                # Display summary table
                if results:
                    st.subheader("Test Results Summary")
                    summary_df = pd.DataFrame(results)
                    st.dataframe(summary_df.style.format({
                        'Max Force (N)': '{:.2f}',
                        'Bending Strength (N/cm²)': '{:.2f}'
                    }))

                    # Generate combined Excel report
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # Write summary sheet
                        summary_df.to_excel(writer, sheet_name='Summary', index=False)

                        # Write raw data for each file
                        for filename, df in dfs:
                            safe_sheet_name = re.sub(r'[\[\]:*?/\\]', '_', filename)[:31]
                            df.to_excel(writer, sheet_name=safe_sheet_name)

                        # Formatting
                        workbook = writer.book
                        summary_sheet = writer.sheets['Summary']
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

                        for col_num, value in enumerate(summary_df.columns):
                            summary_sheet.write(0, col_num, value, header_format)

                        status_col = summary_df.columns.get_loc('Status')
                        for row in range(1, len(summary_df) + 1):
                            status_value = summary_df.loc[row - 1, 'Status']
                            if status_value == "✅ Pass":
                                summary_sheet.write(row, status_col, status_value, pass_format)
                            else:
                                summary_sheet.write(row, status_col, status_value, fail_format)

                        for filename, df in dfs:
                            safe_sheet_name = re.sub(r'[\[\]:*?/\\]', '_', filename)[:31]
                            raw_sheet = writer.sheets[safe_sheet_name]
                            raw_sheet.write_row(0, 0, df.columns.values, header_format)

                        summary_sheet.set_column('A:A', 30)
                        summary_sheet.set_column('B:I', 20)
                        for filename, df in dfs:
                            safe_sheet_name = re.sub(r'[\[\]:*?/\\]', '_', filename)[:31]
                            writer.sheets[safe_sheet_name].set_column('A:Z', 15)

                    # Download button
                    st.download_button(
                        label="Download Combined Excel Report",
                        data=output.getvalue(),
                        file_name=f"Brafe_BendTest_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # Display individual results
                    for res in results:
                        st.divider()
                        st.subheader(f"Results for {res['Filename']}")
                        cols = st.columns(4)
                        cols[0].metric("Part ID", res['Part ID'])
                        cols[1].metric("Job No", res['Job No'])
                        cols[2].metric("Dimensions", f"{res['Support Span (mm)']}×{res['Width (mm)']}×{res['Height (mm)']} mm")
                        cols[3].metric("Status", res['Status'])
                        
                        cols = st.columns(3)
                        cols[0].metric("Max Force", f"{res['Max Force (N)']:.2f} N")
                        cols[1].metric("Bending Strength", f"{res['Bending Strength (N/cm²)']:.2f} N/cm²")
                        cols[2].metric("Target Strength", "260 N/cm²")
                        
                        if res['Bending Strength (N/cm²)'] < 260:
                            st.warning("""
                            **Recommendations to Increase Strength:**
                            - Place parts in oven at 140°C for 3 hours
                            - Increase binder amount
                            - Extend rest period before testing
                            - Check for printing defects (layer separation)
                            """)

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
                
                # Generate Excel report in Brafe template format
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Create summary sheet
                    summary_df = pd.DataFrame({
                        'Parameter': ['Test Date', 'Operator', 'Test ID', 
                                      'Method', 'T1 (g)', 'W1 (g)', 'T2 (g)',
                                      'Mass Loss (g)', 'LOI (%)', 'Status'],
                        'Value': [datetime.datetime.now().strftime('%Y-%m-%d'), 
                                  'Operator Name', 'LOI-001',
                                  method, t1, w1, t2,
                                  abs(delta_m), loi, status]
                    })
                    summary_df.to_excel(writer, sheet_name='Test Summary', index=False)
                    
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
                    
                    # Add Brafe logo (commented out to avoid image error)
                    # summary_sheet.insert_image('A1', 'brafe_logo.png', {'x_offset': 10, 'y_offset': 10})
                    
                    # Set column widths
                    summary_sheet.set_column('A:A', 25)
                    summary_sheet.set_column('B:B', 20)
                
                # Download button
                st.download_button(
                    label="Download Excel Report",
                    data=output.getvalue(),
                    file_name=f"Brafe_LOI_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
