import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import base64
from io import BytesIO
import datetime
import re
import os  # Added the missing import

# Load measurement images and Brafe logo
x_img = Image.open('x_measurement.png')
y_img = Image.open('y_measurement.png')
z_img = Image.open('z_measurement.png')
brafe_logo = Image.open('brafe_logo.png')

# App configuration with updated blue theme
st.set_page_config(
    page_title="Brafe Engineering Quality Control",
    page_icon=brafe_logo,
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
    st.image(brafe_logo, width=150)
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
            st.image(x_img, caption="Measure Dimension X (Length)", use_container_width=True)
            st.info("**X-Dimension:**\n- Length direction\n- Nominal: 172 mm")
        with cols[1]:
            st.image(y_img, caption="Measure Dimension Y (Width)", use_container_width=True)
            st.info("**Y-Dimension:**\n- Width direction\n- Nominal: 22.4 mm")
        with cols[2]:
            st.image(z_img, caption="Measure Dimension Z (Height)", use_container_width=True)
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
        col1, col2, col3 = st.columns(3)
        with col1:
            support_span = st.number_input("Support Span (L) in mm", 
                                          min_value=10.0, 
                                          max_value=200.0, 
                                          value=100.0,
                                          help="Distance between support rods")
        with col2:
            width = st.number_input("Width (b) in mm", 
                                   min_value=10.0, 
                                   max_value=50.0, 
                                   value=22.4,
                                   help="Test bar width dimension")
        with col3:
            height = st.number_input("Height (h) in mm", 
                                    min_value=10.0, 
                                    max_value=50.0, 
                                    value=22.4,
                                    help="Test bar height dimension")
    
    st.divider()
    
    st.subheader("Upload Bend Test Data")
    bend_file = st.file_uploader("Upload CSV from bend test machine", 
                                type=["csv"],
                                help="Should contain force measurements over time")
    
    if bend_file is not None:
        try:
            # Extract part ID and job number from filename (Ben's method)
            filename = bend_file.name
            path_split_ext = os.path.splitext(filename)
            path_split = os.path.split(path_split_ext[0])
            test_key = path_split[1][16:] if len(path_split[1]) >= 16 else path_split[1]
            
            try:
                part_id = test_key.split("(")[0]
                job_no = test_key.split('(')[1][:-1]  # Remove closing parenthesis
            except:
                part_id = "Unknown"
                job_no = "N/A"
            
            # Read CSV with Ben's fixed column names
            df = pd.read_csv(bend_file, 
                             names=['force', 'point_index', 'position', 'time', 'x_axis_measure', 'y_axis_measure'],
                             index_col=False)
            df = df.set_index('time')
            
            # Apply Ben's stress calculations
            df['stress_mpa'] = (3 * df.y_axis_measure * support_span) / (2 * width * height**2)
            df['stress_ncm2'] = df.stress_mpa * 100
            
            # Get max values
            max_force = df['force'].abs().max()
            max_stress_ncm2 = df['stress_ncm2'].max()
            
            status = "✅ Pass" if max_stress_ncm2 >= 260 else "❌ Fail"
            
            # Display results
            col1, col2, col3 = st.columns(3)
            col1.metric("Maximum Force", f"{max_force:.2f} N")
            col2.metric("Bending Strength", f"{max_stress_ncm2:.2f} N/cm²")
            col3.metric("Quality Status", status, 
                       delta=f"Target: 260 N/cm²", 
                       delta_color="normal")
            
            # Show extracted metadata
            st.info(f"**Part ID:** {part_id} | **Job No:** {job_no}")
            
            # Create dual-axis plot for force and stress progression
            fig, ax1 = plt.subplots(figsize=(10, 6))
            
            # Force plot (left axis)
            color = 'tab:blue'
            ax1.set_xlabel('Time (s)')
            ax1.set_ylabel('Force (N)', color=color)
            ax1.plot(df.index, df['force'].abs(), color=color, label='Force')
            ax1.axhline(y=max_force, color=color, linestyle='--', label='Max Force')
            ax1.tick_params(axis='y', labelcolor=color)
            ax1.grid(True)
            
            # Stress plot (right axis)
            ax2 = ax1.twinx()
            color = 'tab:red'
            ax2.set_ylabel('Stress (N/cm²)', color=color)
            ax2.plot(df.index, df['stress_ncm2'], color=color, label='Stress')
            ax2.axhline(y=max_stress_ncm2, color=color, linestyle='--', label='Max Stress')
            ax2.tick_params(axis='y', labelcolor=color)
            
            # Add legends
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax2.legend(lines1 + lines2, labels1 + labels2, loc='upper right')
            
            plt.title('Force and Stress Progression During Bend Test')
            st.pyplot(fig)
            
            # Generate Excel report in Brafe template format
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Create summary sheet
                summary_df = pd.DataFrame({
                    'Parameter': ['Test Date', 'Operator', 'Test ID', 'Part ID', 'Job No',
                                 'Support Span (mm)', 'Width (mm)', 'Height (mm)',
                                 'Max Force (N)', 'Bending Strength (N/cm²)', 'Status'],
                    'Value': [datetime.datetime.now().strftime('%Y-%m-%d'), 
                             'Operator Name', 'TEST-001', part_id, job_no,
                             support_span, width, height,
                             max_force, max_stress_ncm2, status]
                })
                summary_df.to_excel(writer, sheet_name='Test Summary', index=False)
                
                # Create data sheet
                df.to_excel(writer, sheet_name='Raw Data')
                
                # Formatting
                workbook = writer.book
                
                # Summary sheet formatting
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
                
                # Raw data sheet formatting
                raw_sheet = writer.sheets['Raw Data']
                for col_num, value in enumerate(df.reset_index().columns.values):
                    raw_sheet.write(0, col_num, value, header_format)
                
                # Set column widths
                summary_sheet.set_column('A:A', 25)
                summary_sheet.set_column('B:B', 20)
                raw_sheet.set_column('A:Z', 15)
            
            # Download button
            st.download_button(
                label="Download Excel Report",
                data=output.getvalue(),
                file_name=f"Brafe_BendTest_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            if max_stress_ncm2 < 260:
                st.warning("""
                **Recommendations to Increase Strength:**
                - Place parts in oven at 140°C for 3 hours
                - Increase binder amount
                - Extend rest period before testing
                """)
                    
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.info("""
            **Required CSV Format:**
            - Must have exactly 6 columns
            - Columns must contain (in order):
              1. Force measurements
              2. Point index
              3. Position
              4. Time
              5. X-axis measure
              6. Y-axis measure
              
            **Filename Format:**
            - Must follow pattern: `..._XXXXXXXX(JJJ).csv`
            - Where `XXXXXXXX` = Part ID
            - Where `JJJ` = Job Number
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
                    
                    # Add Brafe logo
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
