import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import base64
from io import BytesIO

# Load measurement images and Brafe logo
x_img = Image.open('x_measurement.png')
y_img = Image.open('y_measurement.png')
z_img = Image.open('z_measurement.png')
brafe_logo = Image.open('brafe_logo.png')  # Assuming the uploaded image is saved as brafe_logo.png

# App configuration with Brafe theme
st.set_page_config(
    page_title="Brafe Engineering Quality Control",
    page_icon=brafe_logo,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional Brafe theme and animations
st.markdown(
    """
    <style>
    .stApp {
        background-color: #1a2a44;
        color: #ffffff;
    }
    .stHeader {
        background-color: #d81e5b;
        padding: 10px;
        border-radius: 5px;
        animation: fadeIn 1s ease-in;
    }
    .stTabs [data-baseweb="tab-list"] {
        background-color: #2e4057;
        border-radius: 5px;
    }
    .stTabs [data-baseweb="tab"] {
        color: #ffffff;
        background-color: #2e4057;
        transition: background-color 0.3s;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #3c5c76;
    }
    .stTabs [data-baseweb="tab--selected"] {
        background-color: #d81e5b;
        color: #ffffff;
    }
    .fail-metric {
        background-color: #ff4d4d;
        padding: 5px;
        border-radius: 3px;
        animation: pulse 1s infinite;
    }
    @keyframes fadeIn {
        from { opacity: 0; }
        to { opacity: 1; }
    }
    @keyframes pulse {
        0% { transform: scale(1); }
        50% { transform: scale(1.05); }
        100% { transform: scale(1); }
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

# Main app with Brafe branding
st.markdown('<div class="stHeader"><h1>Brafe Engineering Quality Control Dashboard</h1></div>', unsafe_allow_html=True)
st.subheader("PDB Process Quality Inspection for Printed Parts")

tab1, tab2, tab3 = st.tabs(["Dimensional Check", "3-Point Bend Test", "Loss on Ignition (LOI)"])

with tab1:
    st.header("Dimensional Measurement Verification")
    st.caption("Verify test bar dimensions according to section 3.3 of Quality Control Manual")
    
    with st.expander("📏 Measurement Instructions", expanded=True):
        cols = st.columns(3)
        with cols[0]:
            st.image(x_img, caption="Measure Dimension X (Length)", use_column_width=True)
            st.info("**X-Dimension:**\n- Length direction\n- Nominal: 172 mm")
        with cols[1]:
            st.image(y_img, caption="Measure Dimension Y (Width)", use_column_width=True)
            st.info("**Y-Dimension:**\n- Width direction\n- Nominal: 22.4 mm")
        with cols[2]:
            st.image(z_img, caption="Measure Dimension Z (Height)", use_column_width=True)
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
            
            x_status = "✅ Pass" if abs(x_measured - nominal_x) <= tolerance else "❌ Fail"
            y_status = "✅ Pass" if abs(y_measured - nominal_y) <= tolerance else "❌ Fail"
            z_status = "✅ Pass" if abs(z_measured - nominal_z) <= tolerance else "❌ Fail"
            
            st.subheader("Verification Results")
            cols = st.columns(3)
            with cols[0]:
                st.metric("X-Dimension", f"{x_measured:.1f} mm", 
                         delta=f"Target: {nominal_x}±{tolerance} mm",
                         delta_color="normal" if x_status == "✅ Pass" else "inverse")
                st.markdown(f'<div class="fail-metric" style="display: {"" if x_status == "❌ Fail" else "none"}">{x_status}</div>', unsafe_allow_html=True)
            with cols[1]:
                st.metric("Y-Dimension", f"{y_measured:.1f} mm", 
                         delta=f"Target: {nominal_y}±{tolerance} mm",
                         delta_color="normal" if y_status == "✅ Pass" else "inverse")
                st.markdown(f'<div class="fail-metric" style="display: {"" if y_status == "❌ Fail" else "none"}">{y_status}</div>', unsafe_allow_html=True)
            with cols[2]:
                st.metric("Z-Dimension", f"{z_measured:.1f} mm", 
                         delta=f"Target: {nominal_z}±{tolerance} mm",
                         delta_color="normal" if z_status == "✅ Pass" else "inverse")
                st.markdown(f'<div class="fail-metric" style="display: {"" if z_status == "❌ Fail" else "none"}">{z_status}</div>', unsafe_allow_html=True)
            
            if "✅ Pass" in [x_status, y_status, z_status]:
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
    
    if bend_file:
        try:
            df = pd.read_csv(bend_file, header=None, names=['force', 'point_index', 'position', 'time', 'x_axis', 'y_axis'])
            df['abs_force'] = df['force'].abs()
            max_force = df['abs_force'].max()
            
            bending_strength_mpa = (3 * max_force * support_span) / (2 * width * height**2)
            bending_strength_ncm2 = bending_strength_mpa * 100
            
            status = "✅ Pass" if bending_strength_ncm2 >= 260 else "❌ Fail"
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Maximum Force", f"{max_force:.2f} N")
            col2.metric("Bending Strength", f"{bending_strength_ncm2:.2f} N/cm²")
            col3.metric("Quality Status", status, 
                       delta=f"Target: 260 N/cm²", 
                       delta_color="normal")
            
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(df['time'], df['abs_force'], 'b-', label='Force')
            ax.axhline(y=max_force, color='r', linestyle='--', label='Max Force')
            ax.set_title('Force Progression During Bend Test')
            ax.set_xlabel('Time (s)')
            ax.set_ylabel('Force (N)')
            ax.grid(True)
            ax.legend()
            
            st.pyplot(fig)
            
            with st.expander("View Raw Data"):
                st.dataframe(df)
            
            if bending_strength_ncm2 < 260:
                st.warning("""
                **Recommendations to Increase Strength:**
                - Place parts in oven at 140°C for 3 hours
                - Increase binder amount
                - Extend rest period before testing
                """)
            
            # Export to Excel with color-coded failed readings
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='BendTestData', index=False)
                workbook = writer.book
                worksheet = writer.sheets['BendTestData']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#d81e5b', 'font_color': '#ffffff'})
                fail_format = workbook.add_format({'bg_color': '#ff4d4d', 'font_color': '#ffffff'})
                for col_num, value in enumerate(df.columns):
                    worksheet.write(0, col_num, value, header_format)
                if bending_strength_ncm2 < 260:
                    worksheet.write(1, df.columns.get_loc('abs_force'), max_force, fail_format)
                worksheet.insert_image('A1', 'brafe_logo.png', {'x_offset': 10, 'y_offset': 10})
            st.download_button(
                label="Download Excel Report",
                data=output.getvalue(),
                file_name="Brafe_BendTest_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

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
                
                # Export to Excel with color-coded failed readings
                output = BytesIO()
                data = {'T1': [t1], 'W1': [w1], 'T2': [t2], 'Δm': [abs(delta_m)], 'LOI (%)': [loi]}
                df_loi = pd.DataFrame(data)
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_loi.to_excel(writer, sheet_name='LOI_Data', index=False)
                    workbook = writer.book
                    worksheet = writer.sheets['LOI_Data']
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#d81e5b', 'font_color': '#ffffff'})
                    fail_format = workbook.add_format({'bg_color': '#ff4d4d', 'font_color': '#ffffff'})
                    for col_num, value in enumerate(df_loi.columns):
                        worksheet.write(0, col_num, value, header_format)
                    if loi < 0.5 or loi > 2.5:
                        worksheet.write(1, df_loi.columns.get_loc('LOI (%)'), loi, fail_format)
                    worksheet.insert_image('A1', 'brafe_logo.png', {'x_offset': 10, 'y_offset': 10})
                st.download_button(
                    label="Download Excel Report",
                    data=output.getvalue(),
                    file_name="Brafe_LOI_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
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
**For support:** support@brafeengineering.com | +44 123 456 7890
""")
