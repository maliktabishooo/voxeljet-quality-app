import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image

# Load measurement images
x_img = Image.open('x_measurement.png')
y_img = Image.open('y_measurement.png')
z_img = Image.open('z_measurement.png')

# App configuration
st.set_page_config(
    page_title="Voxeljet Quality Control",
    page_icon="🔬",
    layout="wide"
)

# Sidebar with documentation links
with st.sidebar:
    st.header("Voxeljet Resources")
    st.markdown("[Quality Control Manual](https://www.voxeljet.com/support/)")
    st.markdown("[Technical Support](mailto:service@voxeljet.de)")
    st.markdown("Hotline: +49 821 7483-580")
    
    st.divider()
    st.caption("Test Specifications")
    st.metric("Nominal Bending Strength", "260 N/cm²")
    st.metric("Test Bar Dimensions", "172 × 22.4 × 22.4 mm")
    st.metric("LOI Sample Weight", "30 g")
    st.metric("Dimensional Tolerance", "±0.45 mm")

# Main app
st.title("🔬 Voxeljet Quality Control Dashboard")
st.subheader("PDB Process Quality Inspection for Printed Parts")

tab1, tab2, tab3 = st.tabs(["Dimensional Check", "3-Point Bend Test", "Loss on Ignition (LOI)"])

with tab1:
    st.header("Dimensional Measurement Verification")
    st.caption("Verify test bar dimensions according to section 3.3 of Quality Control Manual")
    
    # Measurement instructions with images
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
    
    # Dimension input form
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
            # Define nominal values and tolerances
            nominal_x = 172.0
            nominal_y = 22.4
            nominal_z = 22.4
            tolerance = 0.45
            
            # Check each dimension
            x_status = "✅ Pass" if abs(x_measured - nominal_x) <= tolerance else "❌ Fail"
            y_status = "✅ Pass" if abs(y_measured - nominal_y) <= tolerance else "❌ Fail"
            z_status = "✅ Pass" if abs(z_measured - nominal_z) <= tolerance else "❌ Fail"
            
            # Display results
            st.subheader("Verification Results")
            cols = st.columns(3)
            with cols[0]:
                st.metric("X-Dimension", f"{x_measured:.1f} mm", 
                         delta=f"Target: {nominal_x}±{tolerance} mm",
                         delta_color="normal" if x_status == "✅ Pass" else "inverse")
                st.info(x_status)
            with cols[1]:
                st.metric("Y-Dimension", f"{y_measured:.1f} mm", 
                         delta=f"Target: {nominal_y}±{tolerance} mm",
                         delta_color="normal" if y_status == "✅ Pass" else "inverse")
                st.info(y_status)
            with cols[2]:
                st.metric("Z-Dimension", f"{z_measured:.1f} mm", 
                         delta=f"Target: {nominal_z}±{tolerance} mm",
                         delta_color="normal" if z_status == "✅ Pass" else "inverse")
                st.info(z_status)
            
            # Overall status
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
            # Read and process CSV data
            df = pd.read_csv(bend_file, header=None, names=['force', 'point_index', 'position', 'time', 'x_axis', 'y_axis'])
            df['abs_force'] = df['force'].abs()
            max_force = df['abs_force'].max()
            
            # Calculate bending strength
            bending_strength_mpa = (3 * max_force * support_span) / (2 * width * height**2)
            bending_strength_ncm2 = bending_strength_mpa * 100  # Convert to N/cm²
            
            # Display results
            col1, col2, col3 = st.columns(3)
            col1.metric("Maximum Force", f"{max_force:.2f} N")
            col2.metric("Bending Strength", f"{bending_strength_ncm2:.2f} N/cm²")
            
            # Quality assessment
            status = "✅ Pass" if bending_strength_ncm2 >= 260 else "❌ Fail"
            col3.metric("Quality Status", status, 
                       delta=f"Target: 260 N/cm²", 
                       delta_color="normal")
            
            # Force vs Time plot
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(df['time'], df['abs_force'], 'b-', label='Force')
            ax.axhline(y=max_force, color='r', linestyle='--', label='Max Force')
            ax.set_title('Force Progression During Bend Test')
            ax.set_xlabel('Time (s)')
            ax.set_ylabel('Force (N)')
            ax.grid(True)
            ax.legend()
            
            st.pyplot(fig)
            
            # Show raw data
            with st.expander("View Raw Data"):
                st.dataframe(df)
                
            # Post-processing recommendations
            if bending_strength_ncm2 < 260:
                st.warning("""
                **Recommendations to Increase Strength:**
                - Place parts in oven at 140°C for 3 hours
                - Increase binder amount
                - Extend rest period before testing
                """)
                
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
                # Calculate LOI
                delta_m = (abs(t2) - abs(t1)) - w1
                loi = (abs(delta_m) / w1) * 100
                
                # Display results
                st.divider()
                st.subheader("Results")
                
                col1, col2 = st.columns(2)
                col1.metric("Mass Loss (Δm)", f"{abs(delta_m):.3f} g")
                col2.metric("Loss on Ignition", f"{loi:.2f} %")
                
                # Interpretation
                st.info("""
                **Interpretation Guide:**
                - Optimal range: 0.5-2.5%
                - < 0.5%: Insufficient binder
                - > 2.5%: Excessive binder
                """)
                
                # Method-specific notes
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

# Footer
st.divider()
st.caption("""
**Quality Control Manual Reference:** PDB_02P06PDBQL2 (Version 0001, Dec 2022)  
**For support:** service@voxeljet.de | +49 821 7483-580
""")
