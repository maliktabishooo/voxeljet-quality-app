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

    with st.expander("⚙️ Test Parameters", expanded=True):
        # Hardcode test bar dimensions as per the manual
        support_span = 172.0
        width = 22.4
        height = 22.4

        st.write(f"**Support Span (L):** {support_span} mm")
        st.write(f"**Width (b):** {width} mm")
        st.write(f"**Height (h):** {height} mm")

    st.divider()

    st.subheader("Upload Bend Test Data")
    bend_files = st.file_uploader("Upload CSV files from bend test machine",
                                    type=["csv"],
                                    accept_multiple_files=True,
                                    help="Should contain force measurements in column 1 (kN)")

    if bend_files:
        # Initialize lists to store results for summary table
        results = []
        dfs = []

        for bend_file in bend_files:
            try:
                # Extract part ID and job number from filename
                filename = bend_file.name
                path_split_ext = os.path.splitext(filename)
                path_split = os.path.split(path_split_ext[0])
                test_key = path_split[1][16:] if len(path_split[1]) >= 16 else path_split[1]

                try:
                    part_id = test_key.split("(")[0]
                    job_no = test_key.split('(')[1][:-1]
                except:
                    part_id = "Unknown"
                    job_no = "N/A"

                # Read CSV, specifying column names to match the file structure
                df = pd.read_csv(bend_file,
                                    names=['force_kn', 'point_index', 'position_mm', 'time', 'x_axis_measure', 'y_axis_measure'],
                                    index_col=False)
                
                # --- START OF CORRECTED CODE ---
                
                # The raw data has a large negative spike at the beginning. 
                # We need to find the point where the actual test begins, i.e.,
                # where the force values are near zero and then start to increase.
                # A robust way is to find the first index where force is close to zero
                # and then find the max force from that point onwards.
                
                # Use the 'y_axis_measure' column as it represents the positive force values
                # Find the first index where the force is below a small threshold (e.g., 1.0 kN)
                start_of_test_index = df[df['y_axis_measure'].abs() < 1.0].first_valid_index()
                
                if start_of_test_index is None:
                    # If no low values are found, assume the test starts at the beginning
                    start_of_test_index = 0
                else:
                    # To be safe, skip a few more rows after the first low value
                    start_of_test_index += 5 

                # Calculate max force from the start of the test onwards
                max_force_kn = df['y_axis_measure'].iloc[start_of_test_index:].max()
                
                # If max_force_kn is still very low or nan, fall back to a simpler method
                if pd.isna(max_force_kn) or max_force_kn < 0.001:
                    max_force_kn = df['y_axis_measure'].max()

                # Convert from kilonewtons (kN) to Newtons (N) for the formula
                max_force_n = max_force_kn * 1000

                # Calculate bending strength: σ = (3FL)/(2bh²)
                bending_strength_nmm2 = (3 * max_force_n * support_span) / (2 * width * height**2)
                bending_strength_ncm2 = bending_strength_nmm2 * 100

                # --- END OF CORRECTED CODE ---
                
                status = "✅ Pass" if bending_strength_ncm2 >= 260 else "❌ Fail"

                # Store results
                results.append({
                    'Filename': filename,
                    'Part ID': part_id,
                    'Job No': job_no,
                    'Max Force (N)': max_force_n,
                    'Bending Strength (N/cm^2)': bending_strength_ncm2,
                    'Status': status
                })
                dfs.append((filename, df))

                # Display individual results
                st.subheader(f"Results for {filename}")
                st.info(f"**Part ID:** {part_id} | **Job No:** {job_no}")
                col1, col2, col3 = st.columns(3)
                col1.metric("Maximum Force", f"{max_force_n:.2f} N")
                col2.metric("Bending Strength", f"{bending_strength_ncm2:.2f} N/cm²")
                col3.metric("Quality Status", status,
                                delta=f"Target: 260 N/cm²",
                                delta_color="normal")

                # Create dual-axis plot for force and stress progression
                # You'll need to define force_n and stress_ncm2 to be able to plot this
                # I've added the definitions for these above.
                fig, ax1 = plt.subplots(figsize=(10, 6))

                # Force plot (left axis)
                color = 'tab:blue'
                ax1.set_xlabel('Time (s)')
                ax1.set_ylabel('Force (N)', color=color)
                ax1.plot(df['time'], df['y_axis_measure'] * 1000, color=color, label='Force')
                ax1.axhline(y=max_force_n, color=color, linestyle='--', label='Max Force')
                ax1.tick_params(axis='y', labelcolor=color)
                ax1.grid(True)

                # Stress plot (right axis)
                ax2 = ax1.twinx()
                color = 'tab:red'
                ax2.set_ylabel('Stress (N/cm²)', color=color)
                
                # Recalculate stress for the plot using the new force data
                df['stress_ncm2'] = ((3 * (df['y_axis_measure'] * 1000) * support_span) / (2 * width * height**2)) * 100
                ax2.plot(df['time'], df['stress_ncm2'], color=color, label='Stress')
                ax2.axhline(y=bending_strength_ncm2, color=color, linestyle='--', label='Max Stress')
                ax2.tick_params(axis='y', labelcolor=color)

                # Add legends
                lines1, labels1 = ax1.get_legend_handles_labels()
                lines2, labels2 = ax2.get_legend_handles_labels()
                ax2.legend(lines1 + lines2, labels1 + labels2, loc='upper right')

                plt.title(f'Force and Stress Progression for {filename}')
                st.pyplot(fig)

                if bending_strength_ncm2 < 260:
                    st.warning("""
                    **Recommendations to Increase Strength:**
                    - Place parts in oven at 140°C for 3 hours
                    - Increase binder amount
                    - Extend rest period before testing
                    """)

            except Exception as e:
                st.error(f"Error processing file {filename}: {str(e)}")
                st.info("""
                **Required CSV Format:**
                - Must have exactly 6 columns
                - Columns must contain (in order):
                  1. Force (kN)
                  2. Point index
                  3. Position (mm)
                  4. Time (s)
                  5. X-axis measure
                  6. Y-axis measure
                """)

        # Display summary table
        if results:
            st.subheader("Summary of All Files")
            summary_df = pd.DataFrame(results)
            st.dataframe(summary_df.style.format({
                'Max Force (N)': '{:.2f}',
                'Bending Strength (N/cm^2)': '{:.2f}'
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

                # Formatting (unchanged)
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
                summary_sheet.set_column('B:F', 20)
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
