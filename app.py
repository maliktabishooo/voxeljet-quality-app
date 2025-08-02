```python
import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import os

# Apply Brafe theme
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
    # st.image("brafe_logo.png", width=150)  # Replace with actual logo path or URL
    st.markdown("**Brafe Logo Placeholder**")  # Temporary placeholder
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
            # st.image("x_img.png", caption="Measure Dimension X (Length)", use_container_width=True)  # Replace with actual image path or URL
            st.markdown("**Image Placeholder: Measure Dimension X (Length)**")  # Temporary placeholder
            st.info("**X-Dimension:**\n- Length direction\n- Nominal: 172 mm")
        with cols[1]:
            # st.image("y_img.png", caption="Measure Dimension Y (Width)", use_container_width=True)  # Replace with actual image path or URL
            st.markdown("**Image Placeholder: Measure Dimension Y (Width)**")  # Temporary placeholder
            st.info("**Y-Dimension:**\n- Width direction\n- Nominal: 22.4 mm")
        with cols[2]:
            # st.image("z_img.png", caption="Measure Dimension Z (Height)", use_container_width=True)  # Replace with actual image path or URL
            st.markdown("**Image Placeholder: Measure Dimension Z (Height)**")  # Temporary placeholder
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
                                help="Should contain force measurements in kN in the third column")
    
    if bend_file is not None:
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
            
            # Read CSV
            df = pd.read_csv(bend_file, 
                             names=['force', 'point_index', 'position', 'time', 'x_axis_measure', 'y_axis_measure'],
                             index_col=False)
            df = df.set_index('time')
            
            # Convert position (assumed kN) to force in Newtons
            conversion_factor = 1000  # kN to N
            df['force_n'] = df['position'] * conversion_factor
            
            # Calculate bending strength: σ = (3FL)/(2bh²)
            df['stress_mpa'] = (3 * df['force_n'] * support_span) / (2 * width * height**2)
            df['stress_ncm2'] = df['stress_mpa'] * 100
            
            # Get max values
            max_force = df['force_n'].abs().max()
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
            
            # Create Chart.js plot
            ```chartjs
            {
                "type": "line",
                "data": {
                    "labels": ${df.index.tolist()},
                    "datasets": [
                        {
                            "label": "Force (N)",
                            "data": ${df['force_n'].abs().tolist()},
                            "borderColor": "#00509d",
                            "backgroundColor": "#00509d",
                            "yAxisID": "y1",
                            "fill": false
                        },
                        {
                            "label": "Stress (N/cm²)",
                            "data": ${df['stress_ncm2'].tolist()},
                            "borderColor": "#cc0000",
                            "backgroundColor": "#cc0000",
                            "yAxisID": "y2",
                            "fill": false
                        }
                    ]
                },
                "options": {
                    "responsive": true,
                    "plugins": {
                        "title": {
                            "display": true,
                            "text": "Force and Stress Progression During Bend Test",
                            "color": "#003366",
                            "font": {
                                "size": 16
                            }
                        },
                        "legend": {
                            "position": "top",
                            "labels": {
                                "color": "#003366"
                            }
                        }
                    },
                    "scales": {
                        "x": {
                            "title": {
                                "display": true,
                                "text": "Time (s)",
                                "color": "#003366"
                            },
                            "grid": {
                                "color": "#a3c6f0"
                            },
                            "ticks": {
                                "color": "#003366"
                            }
                        },
                        "y1": {
                            "type": "linear",
                            "position": "left",
                            "title": {
                                "display": true,
                                "text": "Force (N)",
                                "color": "#00509d"
                            },
                            "grid": {
                                "color": "#a3c6f0"
                            },
                            "ticks": {
                                "color": "#00509d"
                            }
                        },
                        "y2": {
                            "type": "linear",
                            "position": "right",
                            "title": {
                                "display": true,
                                "text": "Stress (N/cm²)",
                                "color": "#cc0000"
                            },
                            "grid": {
                                "display": false
                            },
                            "ticks": {
                                "color": "#cc0000"
                            }
                        }
                    }
                }
            }
            ```
            
            # Generate Excel report
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
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
                df.to_excel(writer, sheet_name='Raw Data')
                
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
                
                raw_sheet = writer.sheets['Raw Data']
                for col_num, value in enumerate(df.reset_index().columns.values):
                    raw_sheet.write(0, col_num, value, header_format)
                
                summary_sheet.set_column('A:A', 25)
                summary_sheet.set_column('B:B', 20)
                raw_sheet.set_column('A:Z', 15)
            
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
              1. Force (unused)
              2. Point index
              3. Position (force in kN)
              4. Time
              5. X-axis measure
              6. Y-axis measure
              
            **Filename Format:**
            - Must follow pattern: ..._XXXXXXXX(JJJ).csv
            - Where XXXXXXXX = Part ID
            - Where JJJ = Job Number
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
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
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
                    
                    summary_sheet.set_column('A:A', 25)
                    summary_sheet.set_column('B:B', 20)
                
                st.download_button(
                    label="Download Excel Report",
                    data=output.getvalue(),
                    file_name=f"Brafe_LOI_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-offic
