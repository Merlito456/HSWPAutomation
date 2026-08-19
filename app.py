import streamlit as st
import pandas as pd
from datetime import datetime
import io
import base64
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import os
import tempfile

# Set page configuration
st.set_page_config(
    page_title="HSWP Automation Tool",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Health and Safety Work Permit Automation")
st.markdown("---")

# Initialize session state for form data
if 'form_data' not in st.session_state:
    st.session_state.form_data = {}

def create_excel_template(data):
    """Create Excel file with filled data"""
    # Check if template exists
    template_path = 'HSWP_template.xlsx'
    if not os.path.exists(template_path):
        st.error("⚠️ Template file 'HSWP_template.xlsx' not found!")
        return None
    
    try:
        # Load the template
        wb = load_workbook(template_path)
        ws = wb.worksheets[0]  # First sheet
        
        # Fill PROJECT DETAILS
        ws['B2'] = data.get('sub_contractor', '')
        ws['D2'] = data.get('requesting_vendor', '')
        ws['B4'] = data.get('project_in_charge', '')
        ws['B5'] = data.get('safety_officer', '')
        ws['B6'] = data.get('project_name', '')
        ws['B7'] = data.get('work_location', '')
        ws['B8'] = data.get('tower_type', '')
        ws['D4'] = data.get('person_in_charge', '')
        ws['D5'] = data.get('work_schedule', '')
        ws['D6'] = data.get('start_date', '')
        ws['D7'] = data.get('end_date', '')
        ws['F5'] = data.get('work_time_period', '')
        ws['F6'] = data.get('start_time', '')
        ws['F7'] = data.get('end_time', '')
        ws['E8'] = data.get('brief_description', '')
        
        # Fill JHA Assessment (top section)
        ws['G3'] = data.get('jha_step1', '')
        ws['H3'] = data.get('jha_hazard1', '')
        ws['I3'] = data.get('jha_control1', '')
        ws['G5'] = data.get('jha_step2', '')
        ws['H5'] = data.get('jha_hazard2', '')
        ws['I5'] = data.get('jha_control2', '')
        
        # Fill WORK DETAILS based on high risk
        high_risk = data.get('high_risk', 'NO')
        if high_risk == 'YES':
            # Work at Heights
            if data.get('work_at_heights', False):
                ws['C12'] = 'X'
            if data.get('scaffold', False):
                ws['D12'] = 'X'
            if data.get('ladder', False):
                ws['E12'] = 'X'
            if data.get('tower', False):
                ws['F12'] = 'X'
            
            # Certifications
            ws['C13'] = data.get('scaffold_cert', '')
            ws['E13'] = data.get('wah_rigger_cert', '')
            if data.get('scaffold_components', False):
                ws['C14'] = 'X'
            if data.get('workers_fit', False):
                ws['E14'] = 'X'
            
            # Electrical Works
            if data.get('electrical_works', False):
                ws['C16'] = 'X'
            ws['C17'] = data.get('electrician_cert', '')
            if data.get('loto_device', False):
                ws['C18'] = 'X'
            if data.get('insulated_tools', False):
                ws['E18'] = 'X'
            
            # Heavy Lifting
            ws['C20'] = data.get('operator_cert', '')
            ws['D20'] = data.get('rigger_cert', '')
            ws['C21'] = data.get('heavy_eqpt_cert', '')
            
            # Confined Space
            if data.get('confined_space', False):
                ws['C23'] = 'X'
            ws['C24'] = data.get('scba_cert', '')
            ws['D24'] = data.get('ventilation_eqpt', '')
            if data.get('flash_arrester', False):
                ws['C25'] = 'X'
            if data.get('fire_blanket', False):
                ws['E25'] = 'X'
            ws['C26'] = data.get('o2_detector', '')
            ws['D26'] = data.get('safety_line', '')
            
            # Harmful Substances
            harmful = data.get('harmful_substance', 'NO')
            ws['B28'] = harmful
            if harmful == 'YES':
                if data.get('fumes', False):
                    ws['C29'] = 'X'
                if data.get('odors', False):
                    ws['D29'] = 'X'
                if data.get('dust', False):
                    ws['C30'] = 'X'
                if data.get('noise', False):
                    ws['D30'] = 'X'
                if data.get('sparks', False):
                    ws['C31'] = 'X'
                ws['D31'] = data.get('other_harmful', '')
            
            # Utility Interruption
            utility = data.get('utility_interruption', 'NO')
            ws['B33'] = utility
            if utility == 'YES':
                ws['C34'] = data.get('affected_utilities', '')
        
        # Fill JHA Table
        jha_steps = data.get('jha_steps', [])
        row_start = 48
        for i, step in enumerate(jha_steps):
            if i >= 10:
                break
            ws[f'A{row_start + i}'] = step.get('step', '')
            ws[f'B{row_start + i}'] = step.get('hazard', '')
            ws[f'D{row_start + i}'] = step.get('controls', '')
        
        # Fill PPE
        ppe_required = data.get('ppe_required', [])
        ppe_mapping = {
            'Safety Shoes': 'B42',
            'Hardhat': 'C42',
            'Body Harness': 'D42',
            'Gloves': 'E42',
            'Welding Mask': 'B43',
            'N95 Masks': 'C43',
            'Goggles': 'D43'
        }
        for ppe, cell in ppe_mapping.items():
            if ppe in ppe_required:
                ws[cell] = 'X'
        
        if 'Other PPE' in ppe_required:
            ws['E43'] = data.get('other_ppe_text', '')
        
        # Fill Tools and Materials
        tools = data.get('tools_materials', [])
        tool_row_start = 39
        for i, tool in enumerate(tools):
            if i >= 10:
                break
            ws[f'A{tool_row_start + i}'] = tool
        
        # Fill Workers
        workers = data.get('workers', [])
        worker_row_start = 44
        for i, worker in enumerate(workers):
            if i >= 8:
                break
            ws[f'A{worker_row_start + i}'] = worker
        
        # Fill Acknowledgement
        ws['B51'] = data.get('prepared_by', '')
        ws['C51'] = data.get('noted_by', '')
        ws['E51'] = data.get('approved_by', '')
        ws['C52'] = data.get('noted_by', '')
        if data.get('approved_status') == 'YES':
            ws['E52'] = 'X'
        ws['G52'] = data.get('safety_officer_approval', '')
        
        return wb
    
    except Exception as e:
        st.error(f"Error processing template: {str(e)}")
        return None

def get_excel_download_link(wb, filename):
    """Generate download link for Excel file"""
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    b64 = base64.b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" style="background-color: #4CAF50; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; display: inline-block;">📥 Download Completed Excel File</a>'
    return href

def main():
    # Create two columns for layout
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📋 Project Details")
        
        # Project Details
        sub_contractor = st.text_input("Name of Sub Contractor", value="Ultegra Supplies and Services")
        requesting_vendor = st.text_input("Requesting Vendor", value="NOKIA SHANGHAI BELL")
        project_in_charge = st.text_input("Project In-Charge")
        safety_officer = st.text_input("Project Safety Officer", value="MELVIN ADOVE")
        project_name = st.text_input("Project Name", value="FIBERTIME")
        work_location = st.text_input("Work Location", value="Naic, Cavite")
        tower_type = st.text_input("Tower Type", value="roofdeck")
        person_in_charge = st.text_input("Person In-charge")
        work_schedule = st.text_input("Work Schedule")
        
        col1a, col1b = st.columns(2)
        with col1a:
            start_date = st.date_input("Start Date", value=datetime.now().date())
            start_time = st.time_input("Start Time", value=datetime.now().time())
        with col1b:
            end_date = st.date_input("End Date", value=datetime.now().date())
            end_time = st.time_input("End Time", value=datetime.now().time())
        
        brief_description = st.text_area("Brief Description of Work", value="Service testing")
        
        st.subheader("🛠️ Work Details")
        
        # High Risk
        high_risk = st.radio("Is work to be done with high risk?", ["NO", "YES"])
        
        if high_risk == "YES":
            with st.expander("Work at Heights", expanded=True):
                col_wh1, col_wh2 = st.columns(2)
                with col_wh1:
                    work_at_heights = st.checkbox("Work at Heights")
                    scaffold = st.checkbox("Scaffold")
                    scaffold_cert = st.text_input("NCII Cert of Scaffold Erector")
                with col_wh2:
                    ladder = st.checkbox("Ladder")
                    tower = st.checkbox("Tower")
                    wah_rigger_cert = st.text_input("WAH Rigger Certificate")
                
                scaffold_components = st.checkbox("Scaffold components available")
                workers_fit = st.checkbox("Workers physically fit")
            
            with st.expander("Electrical Works", expanded=True):
                electrical_works = st.checkbox("Electrical Works")
                electrician_cert = st.text_input("NCII Cert of Electrician or ID of REE/RME")
                loto_device = st.checkbox("LOTO Device")
                insulated_tools = st.checkbox("Insulated Tools")
            
            with st.expander("Heavy Lifting / Excavation", expanded=True):
                operator_cert = st.text_input("NCII Cert of Operator")
                rigger_cert = st.text_input("Cert of Lift Rigger")
                heavy_eqpt_cert = st.text_input("3rd Party Certification of Heavy Eqpt")
            
            with st.expander("Confined Space Works", expanded=True):
                confined_space = st.checkbox("Confined Space Works")
                col_cs1, col_cs2 = st.columns(2)
                with col_cs1:
                    scba_cert = st.text_input("Certificate of SCBA Operator")
                    ventilation_eqpt = st.text_input("Ventilation Equipment")
                    flash_arrester = st.checkbox("OxyFuel flash back arrester installed")
                with col_cs2:
                    fire_blanket = st.checkbox("Fire Blanket")
                    o2_detector = st.text_input("O2 and Gas Detector")
                    safety_line = st.text_input("Safety Line")
            
            with st.expander("Harmful Substances", expanded=True):
                harmful_substance = st.radio("Is there any harmful substance or nuisance release?", ["NO", "YES"])
                if harmful_substance == "YES":
                    col_h1, col_h2 = st.columns(2)
                    with col_h1:
                        fumes = st.checkbox("Fumes")
                        dust = st.checkbox("Dust")
                        sparks = st.checkbox("Sparks")
                    with col_h2:
                        odors = st.checkbox("Offensive Odors")
                        noise = st.checkbox("Noise")
                        other_harmful = st.text_input("Others:")
            
            with st.expander("Utility Interruption", expanded=True):
                utility_interruption = st.radio("Will there be Utility interruption?", ["NO", "YES", "N/A"])
                if utility_interruption == "YES":
                    affected_utilities = st.text_area("Specify affected utilities and affected areas")
        else:
            # Set defaults for non-high risk
            work_at_heights = scaffold = ladder = tower = False
            scaffold_cert = wah_rigger_cert = ""
            scaffold_components = workers_fit = False
            electrical_works = False
            electrician_cert = loto_device = insulated_tools = ""
            operator_cert = rigger_cert = heavy_eqpt_cert = ""
            confined_space = False
            scba_cert = ventilation_eqpt = flash_arrester = fire_blanket = ""
            o2_detector = safety_line = ""
            harmful_substance = "NO"
            fumes = odors = dust = noise = sparks = False
            other_harmful = ""
            utility_interruption = "NO"
            affected_utilities = ""
    
    with col2:
        st.subheader("⚠️ Job Hazard Assessment")
        
        # Main JHA Table
        st.markdown("**JOB HAZARD ASSESSMENT**")
        
        # Step 1
        jha_step1 = st.text_input("Job Step 1", value="Site Access")
        jha_hazard1 = st.text_input("Hazard 1", value="Slip, trip, and fall from uneven surface")
        jha_control1 = st.text_area("Control 1", value="Coordinate with lessor/UDI security prior to entry. Ensure all permits are on hand. Only authorized person can enter the area", height=60)
        
        # Step 2
        jha_step2 = st.text_input("Job Step 2", value="Prepare Work Area")
        jha_hazard2 = st.text_input("Hazard 2", value="Trips/Falls: Uneven surfaces, debris, inadequate lighting")
        jha_control2 = st.text_area("Control 2", value="Clear work area of debris, ensure adequate lighting, wear appropriate footwear (safety shoes)", height=60)
        
        # Step 3 - Handling Fiber Optics (default)
        jha_step3 = st.text_input("Job Step 3", value="Handling of Fiber Optic cables")
        jha_hazard3 = st.text_input("Hazard 3", value="Cuts/Lacerations: Sharp edges of fibers or connectors.")
        jha_control3 = st.text_area("Control 3", value="Wear cut-resistant gloves. Handle fibers with care", height=60)
        
        # Additional JHA entries
        st.markdown("---")
        st.subheader("Additional JHA Entries")
        
        # Create JHA steps list with defaults
        jha_steps = []
        
        # Add the first three steps
        if jha_step1 and jha_hazard1 and jha_control1:
            jha_steps.append({
                'step': jha_step1,
                'hazard': jha_hazard1,
                'controls': jha_control1
            })
        
        if jha_step2 and jha_hazard2 and jha_control2:
            jha_steps.append({
                'step': jha_step2,
                'hazard': jha_hazard2,
                'controls': jha_control2
            })
        
        if jha_step3 and jha_hazard3 and jha_control3:
            jha_steps.append({
                'step': jha_step3,
                'hazard': jha_hazard3,
                'controls': jha_control3
            })
        
        # Additional JHA rows
        num_extra = st.number_input("Number of additional JHA entries", min_value=0, max_value=7, value=0)
        
        for i in range(num_extra):
            with st.expander(f"JHA Entry {i+4}", expanded=False):
                step = st.text_input(f"Job Step {i+4}", key=f"jha_step_{i}")
                hazard = st.text_input(f"Hazard {i+4}", key=f"jha_hazard_{i}")
                controls = st.text_area(f"Controls {i+4}", key=f"jha_controls_{i}", height=60)
                if step and hazard and controls:
                    jha_steps.append({
                        'step': step,
                        'hazard': hazard,
                        'controls': controls
                    })
        
        st.subheader("🦺 Required PPE")
        ppe_options = ['Safety Shoes', 'Hardhat', 'Body Harness', 'Gloves', 'Welding Mask', 'N95 Masks', 'Goggles', 'Other PPE']
        ppe_required = st.multiselect("Select required PPE", ppe_options)
        other_ppe_text = ""
        if 'Other PPE' in ppe_required:
            other_ppe_text = st.text_input("Other PPE details")
        
        st.subheader("🔧 Tools and Materials")
        tools_text = st.text_area("List tools and materials (one per line)", 
                                  value="PPE\nFIRST AID KIT\nODF\nOPM\nOTDR\nPatchcord")
        tools = [t.strip() for t in tools_text.split('\n') if t.strip()]
        
        st.subheader("👷 List of Workers")
        workers_text = st.text_area("List workers (one per line)",
                                   value="JAY PALASOL\nROSE EISELE BARBA\nEDGAR PERALTA\nMELVIN ADOVE\nROCKY MARZO TRIVIÑO\nJANN ALEXIS AGULO\nRODERICK REYES\nMARK JOSEPH INOVERO\nMARK JAYSON BRILLANTES")
        workers = [w.strip() for w in workers_text.split('\n') if w.strip()]
        
        st.subheader("✅ Acknowledgement")
        prepared_by = st.text_input("Prepared By")
        noted_by = st.text_input("Noted By", value="Ma. Lorena Angelica Tenorio")
        approved_by = st.radio("Approved By", ["YES", "NO"])
        safety_officer_approval = st.text_input("MIDC Safety Officer")

    # Collect all data
    data = {
        'sub_contractor': sub_contractor,
        'requesting_vendor': requesting_vendor,
        'project_in_charge': project_in_charge,
        'safety_officer': safety_officer,
        'project_name': project_name,
        'work_location': work_location,
        'tower_type': tower_type,
        'person_in_charge': person_in_charge,
        'work_schedule': work_schedule,
        'start_date': start_date.strftime('%Y-%m-%d'),
        'end_date': end_date.strftime('%Y-%m-%d'),
        'start_time': start_time.strftime('%H:%M'),
        'end_time': end_time.strftime('%H:%M'),
        'brief_description': brief_description,
        'high_risk': high_risk,
        'work_at_heights': locals().get('work_at_heights', False),
        'scaffold': locals().get('scaffold', False),
        'ladder': locals().get('ladder', False),
        'tower': locals().get('tower', False),
        'scaffold_cert': locals().get('scaffold_cert', ''),
        'wah_rigger_cert': locals().get('wah_rigger_cert', ''),
        'scaffold_components': locals().get('scaffold_components', False),
        'workers_fit': locals().get('workers_fit', False),
        'electrical_works': locals().get('electrical_works', False),
        'electrician_cert': locals().get('electrician_cert', ''),
        'loto_device': locals().get('loto_device', False),
        'insulated_tools': locals().get('insulated_tools', False),
        'operator_cert': locals().get('operator_cert', ''),
        'rigger_cert': locals().get('rigger_cert', ''),
        'heavy_eqpt_cert': locals().get('heavy_eqpt_cert', ''),
        'confined_space': locals().get('confined_space', False),
        'scba_cert': locals().get('scba_cert', ''),
        'ventilation_eqpt': locals().get('ventilation_eqpt', ''),
        'flash_arrester': locals().get('flash_arrester', False),
        'fire_blanket': locals().get('fire_blanket', False),
        'o2_detector': locals().get('o2_detector', ''),
        'safety_line': locals().get('safety_line', ''),
        'harmful_substance': locals().get('harmful_substance', 'NO'),
        'fumes': locals().get('fumes', False),
        'odors': locals().get('odors', False),
        'dust': locals().get('dust', False),
        'noise': locals().get('noise', False),
        'sparks': locals().get('sparks', False),
        'other_harmful': locals().get('other_harmful', ''),
        'utility_interruption': locals().get('utility_interruption', 'NO'),
        'affected_utilities': locals().get('affected_utilities', ''),
        'jha_step1': jha_step1,
        'jha_hazard1': jha_hazard1,
        'jha_control1': jha_control1,
        'jha_step2': jha_step2,
        'jha_hazard2': jha_hazard2,
        'jha_control2': jha_control2,
        'jha_steps': jha_steps,
        'ppe_required': ppe_required,
        'other_ppe_text': other_ppe_text,
        'tools_materials': tools,
        'workers': workers,
        'prepared_by': prepared_by,
        'noted_by': noted_by,
        'approved_by': approved_by,
        'approved_status': approved_by,
        'safety_officer_approval': safety_officer_approval
    }
    
    st.markdown("---")
    
    # Generate Excel button
    if st.button("📥 Generate Excel File", type="primary"):
        with st.spinner("Generating Excel file..."):
            try:
                # Create Excel file
                wb = create_excel_template(data)
                if wb:
                    # Generate download link
                    filename = f"HSWP_{project_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    download_link = get_excel_download_link(wb, filename)
                    
                    st.success("✅ Excel file generated successfully!")
                    st.markdown(download_link, unsafe_allow_html=True)
                    
                    # Preview data
                    with st.expander("📊 Preview Data Summary"):
                        st.json(data)
                else:
                    st.error("❌ Failed to generate Excel file. Please check the template file.")
                    
            except Exception as e:
                st.error(f"❌ Error generating file: {str(e)}")
                st.exception(e)
    
    # Instructions
    with st.sidebar:
        st.header("📝 Instructions")
        st.markdown("""
        1. Fill in all required fields in the main form
        2. Add additional JHA entries as needed
        3. Select required PPE
        4. List tools, materials, and workers
        5. Click 'Generate Excel File' button
        6. Download the completed Excel file
        """)
        
        st.header("📋 Template Requirements")
        st.markdown("""
        - The template file 'HSWP_template.xlsx' must be in the repository
        - All fields marked with * are required
        - The generated file will be named with the project name and timestamp
        """)
        
        st.header("ℹ️ About")
        st.markdown("""
        This tool automates the creation of Health and Safety Work Permits.
        
        **Version:** 1.0.0
        **Last Updated:** 2026-08-19
        """)

if __name__ == "__main__":
    main()
