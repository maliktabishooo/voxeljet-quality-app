import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import base64
from io import BytesIO

# Load measurement images and Brafe logo
x_img = Image.open('x_measurement.png')  # Load image for X-dimension measurement
y_img = Image.open('y_measurement.png')  # Load image for Y-dimension measurement
z_img = Image.open('z_measurement.png')  # Load image for Z-dimension measurement
brafe_logo = Image.open('brafe_logo.png')  # Load Brafe Engineering logo image

# App configuration with Brafe theme
st.set_page_config(
    page_title="Brafe Engineering Quality Control",  # Set the browser tab title
    page_icon=brafe_logo,  # Set the app icon to Brafe logo
    layout="wide",  # Use wide layout for better visibility
    initial_sidebar_state="expanded"  # Expand sidebar by default
)

# Custom CSS for lighter Brafe theme and animations
st.markdown(
    """
    <style>
    .stApp {
        background-color: #e6eef7;  /* Lighter blue background for a softer look */
        color: #1a2a44;  /* Dark blue text for contrast */
    }
    .stHeader {
        background-color: #d81e5b;  /* Brafe red header background */
        padding: 10px;
        border-radius: 5px;
        animation: fadeIn 1s ease-in;  /* Fade-in animation for header */
    }
    .stTabs [data-baseweb="tab-list"] {
        background-color: #b0c4de;  /* Light steel blue for tab background */
        border-radius: 5px;
    }
    .stTabs [data-baseweb="tab"] {
        color: #1a2a44;  /* Dark blue tab text */
        background-color: #b0c4de;  /* Match tab list background */
        transition: background-color 0.3s;  /* Smooth hover effect */
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #a3bffa;  /* Lighter blue on hover */
    }
    .stTabs [data-baseweb="tab--selected"] {
        background-color: #d81e5b;  /* Red for selected tab */
        color: #ffffff;
    }
    .fail-metric {
        background-color: #ff4d4d;  /* Red background for failed metrics */
        padding: 5px;
        border-radius: 3px;
        animation: pulse 1s infinite;  /* Pulse animation for failed metrics */
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
    st.image(brafe_logo, width=150)  # Display Brafe logo in sidebar
    st.header("Brafe Engineering Resources")  # Sidebar header
    st.markdown("[Quality Control Manual](https://www.brafeengineering.com/support/)")  # Link to manual
    st.markdown("[Technical Support](mailto:support@brafeengineering.com)")  # Support email
    st.markdown("Hotline: +44 123 456 7890")  # Contact number
    
    st.divider()  # Visual separator
    st.caption("Test Specifications")  # Section caption
    st.metric("Nominal Bending Strength", "260 N/cm²")  # Display spec
    st.metric("Test Bar Dimensions", "172 × 22.4 × 22.4 mm")  # Display spec
    st.metric("LOI Sample Weight", "30 g")  # Display spec
    st.metric("Dimensional Tolerance", "±0.45 mm")  # Display spec

# Main app with Brafe branding
st.markdown('<div class="stHeader"><h1>Brafe Engineering Quality Control Dashboard</h1></div>', unsafe_allow_html=True)  # Custom header
st.subheader("PDB Process Quality Inspection for Printed Parts")  # Subheader

tab1, tab2, tab3 = st.tabs(["Dimensional Check", "3-Point Bend Test", "Loss on Ignition (LOI)"])  # Create tabbed interface

with tab1:
    st.header("Dimensional Measurement Verification")  # Tab header
    st.caption("Verify test bar dimensions according to section 3.3 of Quality Control Manual")  # Caption
    
    with st.expander("📏 Measurement Instructions", expanded=True):  # Expandable section
        cols = st.columns(3)  # Create three columns
        with cols[0]:
            st.image(x_img, caption="Measure Dimension X (Length)", use_container_width=True)  # Display X image with container width
            st.info("**X-Dimension:**\n- Length direction\n- Nominal: 172 mm")  # Info text
        with cols[1]:
            st.image(y_img, caption="Measure Dimension Y (Width)", use_container_width=True)  # Display Y image with container width
            st.info("**Y-Dimension:**\n- Width direction\n- Nominal: 22.4 mm")  # Info text
        with cols[2]:
            st.image(z_img, caption="Measure Dimension Z (Height)", use_container_width=True)  # Display Z image with container width
            st.info("**Z-Dimension:**\n- Height direction\n- Nominal: 22.4 mm")  # Info text
        
        st.markdown("""
        **Procedure:**
        1. Ensure test bar has rested in sand for ≥6 hours
        2. Clean loose sand from test bar
        3. Measure each dimension with measuring slide
        4. Compare with nominal values (tolerance ±0.45mm)
        """)  # Procedure steps
    
    st.divider()  # Visual separator
    
    with st.form("dimensional_check"):  # Form for dimension input
        st.subheader("Enter Measured Dimensions (mm)")  # Form header
        col1, col2, col3 = st.columns(3)  # Three columns for inputs
        
        with col1:
            x_measured = st.number_input("X-Dimension (Length)", 
                                        min_value=150.0,
                                        max_value=190.0,
                                        value=172.0,
                                        step=0.1,
                                        format="%.1f")  # Input for X dimension
        with col2:
            y_measured = st.number_input("Y-Dimension (Width)", 
                                        min_value=15.0,
                                        max_value=30.0,
                                        value=22.4,
                                        step=0.1,
                                        format="%.1f")  # Input for Y dimension
        with col3:
            z_measured = st.number_input("Z-Dimension (Height)", 
                                        min_value=15.0,
                                        max_value=30.0,
                                        value=22.4,
                                        step=0.1,
                                        format="%.1f")  # Input for Z dimension
        
        submitted = st.form_submit_button("Verify Dimensions")  # Submit button
        
        if submitted:
            nominal_x, nominal_y, nominal_z = 172.0, 22.4, 22.4  # Nominal values
            tolerance = 0.45  # Tolerance value
            
            x_status = "✅ Pass" if abs(x_measured - nominal_x) <= tolerance else "❌ Fail"  # Check X status
            y_status = "✅ Pass" if abs(y_measured - nominal_y) <= tolerance else "❌ Fail"  # Check Y status
            z_status = "✅ Pass" if abs(z_measured - nominal_z) <= tolerance else "❌ Fail"  # Check Z status
            
            st.subheader("Verification Results")  # Results header
            cols = st.columns(3)  # Three columns for results
            with cols[0]:
                st.metric("X-Dimension", f"{x_measured:.1f} mm", 
                         delta=f"Target: {nominal_x}±{tolerance} mm",
                         delta_color="normal" if x_status == "✅ Pass" else "inverse")  # X metric
                st.markdown(f'<div class="fail-metric" style="display: {"" if x_status == "❌ Fail" else "none"}">{x_status}</div>', unsafe_allow_html=True)  # Show fail status
            with cols[1]:
                st.metric("Y-Dimension", f"{y_measured:.1f} mm", 
                         delta=f"Target: {nominal_y}±{tolerance} mm",
                         delta_color="normal" if y_status == "✅ Pass" else "inverse")  # Y metric
                st.markdown(f'<div class="fail-metric" style="display: {"" if y_status == "❌ Fail" else "none"}">{y_status}</div>', unsafe_allow_html=True)  # Show fail status
            with cols[2]:
                st.metric("Z-Dimension", f"{z_measured:.1f} mm", 
                         delta=f"Target: {nominal_z}±{tolerance} mm",
                         delta_color="normal" if z_status == "✅ Pass" else "inverse")  # Z metric
                st.markdown(f'<div class="fail-metric" style="display: {"" if z_status == "❌ Fail" else "none"}">{z_status}</div>', unsafe_allow_html=True)  # Show fail status
            
            if "✅ Pass" in [x_status, y_status, z_status]:  # Check overall pass/fail
                st.success("All dimensions within specification!")
            else:
                st.error("Some dimensions out of tolerance. Check print parameters.")
                st.markdown("""
                **Troubleshooting Tips:**
                - Check offset values in defr3d.ini file
                - Verify zCompensation parameter
                - Ensure proper printer calibration
                """)  # Troubleshooting advice

with tab2:
    st.header("3-Point Bend Test Analysis")  # Tab header
    st.caption("Calculate bending strength according to section 3.4 of Quality Control Manual")  # Caption
    
    with st.expander("⚙️ Test Parameters", expanded=True):  # Expandable section for parameters
        col1, col2, col3 = st.columns(3)  # Three columns for inputs
        with col1:
            support_span = st.number_input("Support Span (L) in mm", 
                                          min_value=10.0, 
                                          max_value=200.0, 
                                          value=100.0,
                                          help="Distance between support rods")  # Input for support span
        with col2:
            width = st.number_input("Width (b) in mm", 
                                   min_value=10.0, 
                                   max_value=50.0, 
                                   value=22.4,
                                   help="Test bar width dimension")  # Input for width
        with col3:
            height = st.number_input("Height (h) in mm", 
                                    min_value=10.0, 
                                    max_value=50.0, 
                                    value=22.4,
                                    help="Test bar height dimension")  # Input for height
    
    st.divider()  # Visual separator
    
    st.subheader("Upload Bend Test Data")  # Subheader for file upload
    bend_file = st.file_uploader("Upload CSV from bend test machine", 
                                type=["csv"],
                                help="Should contain force measurements over time")  # File uploader
    
    if bend_file:
        try:
            df = pd.read_csv(bend_file, header=None, names=['force', 'point_index', 'position', 'time', 'x_axis', 'y_axis'])  # Read CSV data
            df['abs_force'] = df['force'].abs()  # Calculate absolute force
            max_force = df['abs_force'].max()  # Find maximum force
            
            bending_strength_mpa = (3 * max_force * support_span) / (2 * width * height**2)  # Calculate bending strength in MPa
            bending_strength_ncm2 = bending_strength_mpa * 100  # Convert to N/cm²
            
            status = "✅ Pass" if bending_strength_ncm2 >= 260 else "❌ Fail"  # Determine pass/fail status
            
            col1, col2, col3 = st.columns(3)  # Three columns for metrics
            col1.metric("Maximum Force", f"{max_force:.2f} N")  # Display max force
            col2.metric("Bending Strength", f"{bending_strength_ncm2:.2f} N/cm²")  # Display bending strength
            col3.metric("Quality Status", status, 
                       delta=f"Target: 260 N/cm²", 
                       delta_color="normal")  # Display status
            
            fig, ax = plt.subplots(figsize=(10, 4))  # Create plot
            ax.plot(df['time'], df['abs_force'], 'b-', label='Force')  # Plot force over time
            ax.axhline(y=max_force, color='r', linestyle='--', label='Max Force')  # Add max force line
            ax.set_title('Force Progression During Bend Test')  # Set title
            ax.set_xlabel('Time (s)')  # Set x-axis label
            ax.set_ylabel('Force (N)')  # Set y-axis label
            ax.grid(True)  # Add grid
            ax.legend()  # Add legend
            st.pyplot(fig)  # Display plot
            
            with st.expander("View Raw Data"):  # Expandable section for raw data
                st.dataframe(df)  # Display dataframe
            
            if bending_strength_ncm2 < 260:  # Provide recommendations if failed
                st.warning("""
                **Recommendations to Increase Strength:**
                - Place parts in oven at 140°C for 3 hours
                - Increase binder amount
                - Extend rest period before testing
                """)
            
            # Export to Excel with color-coded failed readings
            output = BytesIO()  # Create in-memory file
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='BendTestData', index=False)  # Write data to Excel
                workbook = writer.book
                worksheet = writer.sheets['BendTestData']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#d81e5b', 'font_color': '#ffffff'})  # Header format
                fail_format = workbook.add_format({'bg_color': '#ff4d4d', 'font_color': '#ffffff'})  # Fail format
                for col_num, value in enumerate(df.columns):
                    worksheet.write(0, col_num, value, header_format)  # Write headers
                if bending_strength_ncm2 < 260:
                    worksheet.write(1, df.columns.get_loc('abs_force'), max_force, fail_format)  # Highlight fail
                worksheet.insert_image('A1', 'brafe_logo.png', {'x_offset': 10, 'y_offset': 10})  # Insert logo
            st.download_button(
                label="Download Excel Report",
                data=output.getvalue(),
                file_name="Brafe_BendTest_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )  # Download button
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")  # Error handling

with tab3:
    st.header("Loss on Ignition (LOI) Analysis")  # Tab header
    st.caption("Calculate binder content according to section 3.5 of Quality Control Manual")  # Caption
    
    method = st.radio("Test Method", 
                     ["Bunsen Burner (Section 3.5.1)", "Oven (Section 3.5.2)"],
                     index=0,
                     horizontal=True)  # Radio button for test method
    
    with st.form("loi_calculation"):  # Form for LOI input
        st.subheader("Enter Measurement Values")  # Form header
        col1, col2, col3 = st.columns(3)  # Three columns for inputs
        
        with col1:
            t1 = st.number_input("T1 (Bowl Weight) in g", 
                                value=-44.904,
                                format="%.3f",
                                help="Negative value shown on scale after taring")  # Input for T1
        with col2:
            w1 = st.number_input("W1 (Sample Weight) in g", 
                                min_value=20.0,
                                max_value=40.0,
                                value=30.023,
                                format="%.3f")  # Input for W1
        with col3:
            t2 = st.number_input("T2 (Bowl + Ash) in g", 
                                value=-74.422,
                                format="%.3f",
                                help="Negative value shown on scale after taring")  # Input for T2
        
        submitted = st.form_submit_button("Calculate LOI")  # Submit button
        
        if submitted:
            try:
                delta_m = (abs(t2) - abs(t1)) - w1  # Calculate mass loss
                loi = (abs(delta_m) / w1) * 100  # Calculate LOI percentage
                
                st.divider()  # Visual separator
                st.subheader("Results")  # Results header
                
                col1, col2 = st.columns(2)  # Two columns for metrics
                col1.metric("Mass Loss (Δm)", f"{abs(delta_m):.3f} g")  # Display mass loss
                col2.metric("Loss on Ignition", f"{loi:.2f} %")  # Display LOI
                
                st.info("""
                **Interpretation Guide:**
                - Optimal range: 0.5-2.5%
                - < 0.5%: Insufficient binder
                - > 2.5%: Excessive binder
                """)  # Interpretation guide
                
                if "Bunsen" in method:  # Method-specific notes
                    st.caption("Bunsen Burner Method Notes:\n- Burn until sand turns white\n- Stir every minute\n- Cool for 20 min before weighing")
                else:
                    st.caption("Oven Method Notes:\n- Heat to 900°C for 3 hours\n- Cool in closed oven before weighing")
                
                # Export to Excel with color-coded failed readings
                output = BytesIO()  # Create in-memory file
                data = {'T1': [t1], 'W1': [w1], 'T2': [t2], 'Δm': [abs(delta_m)], 'LOI (%)': [loi]}  # Data for Excel
                df_loi = pd.DataFrame(data)
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_loi.to_excel(writer, sheet_name='LOI_Data', index=False)  # Write data to Excel
                    workbook = writer.book
                    worksheet = writer.sheets['LOI_Data']
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#d81e5b', 'font_color': '#ffffff'})  # Header format
                    fail_format = workbook.add_format({'bg_color': '#ff4d4d', 'font_color': '#ffffff'})  # Fail format
                    for col_num, value in enumerate(df_loi.columns):
                        worksheet.write(0, col_num, value, header_format)  # Write headers
                    if loi < 0.5 or loi > 2.5:
                        worksheet.write(1, df_loi.columns.get_loc('LOI (%)'), loi, fail_format)  # Highlight fail
                    worksheet.insert_image('A1', 'brafe_logo.png', {'x_offset': 10, 'y_offset': 10})  # Insert logo
                st.download_button(
                    label="Download Excel Report",
                    data=output.getvalue(),
                    file_name="Brafe_LOI_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )  # Download button
            except Exception as e:
                st.error(f"Calculation error: {str(e)}")  # Error handling
    
    st.divider()  # Visual separator
    st.subheader("LOI Formula Reference")  # Formula header
    st.latex(r'''
    \begin{align*}
    \Delta m &= (|T2| - |T1|) - W1 \\
    LOI (\%) &= \left( \frac{|\Delta m|}{W1} \right) \times 100
    \end{align*}
    ''')  # Display LaTeX formula
    st.caption("Note: Algebraic signs are not considered in calculations (per manual section 3.5)")  # Formula note

# Footer with Brafe branding
st.divider()  # Visual separator
st.caption("""
**Quality Control Manual Reference:** PDB_02P06PDBQL2 (Version 0001, Dec 2022) 
""")  # Footer information
