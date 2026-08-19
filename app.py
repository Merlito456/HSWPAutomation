import streamlit as st
import pandas as pd
from datetime import datetime
import io
import base64
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import tempfile
import os

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
    # Create a new workbook from template
    wb = load_workbook('HSWP_template.xlsx')
    
    # Select the first sheet (Work Permit)
    ws = wb.worksheets[0]
    
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
    
    # Fill WORK DETAILS
    # High Risk
    high_risk = data.get('high_risk', 'NO')
    ws['B12'] = high_risk
    if high_risk == 'YES':
        # Work at Heights
        ws['C12'] = 'X' if data.get('work_at_heights', False) else ''
        ws['D12'] = 'X' if data.get('scaffold', False) else ''
        ws['E12'] = 'X' if data.get('ladder', False) else ''
        ws['F12'] = 'X' if data.get('tower', False) else ''
        
        # Certifications
        ws['C13'] = data.get('scaffold_cert', '')
        ws['E13'] = data.get('wah_rigger_cert', '')
        ws['C14'] = 'X' if data.get('scaffold_components', False) else ''
        ws['E14'] = 'X' if data.get('workers_fit', False) else ''
        
        # Electrical Works
        ws['C16'] = 'X' if data.get('electrical_works', False) else ''
        ws['C17'] = data.get('electrician_cert', '')
        ws['C18'] = 'X' if data.get('loto_device', False) else ''
        ws['E18'] = 'X' if data.get('insulated_tools', False) else ''
        
        # Heavy Lifting
        ws['C20'] = data.get('operator_cert', '')
        ws['D20'] = data.get('rigger_cert', '')
        ws['C21'] = data.get('heavy_eqpt_cert', '')
        
        # Confined Space
        ws['C23'] = 'X' if data.get('confined_space', False) else ''
        ws['C24'] = data.get('scba_cert', '')
        ws['D24'] = data.get('ventilation_eqpt', '')
        ws['C25'] = 'X' if data.get('flash_arrester', False) else ''
        ws['E25'] = 'X' if data.get('fire_blanket', False) else ''
        ws['C26'] = data.get('o2_detector', '')
        ws['D26'] = data.get('safety_line', '')
        
        # Harmful Substances
        harmful = data.get('harmful_substance', 'NO')
        ws['B28'] = harmful
        if harmful == 'YES':
            ws['C29'] = 'X' if data.get('fumes', False) else ''
            ws['D29'] = 'X' if data.get('odors', False) else ''
            ws['C30'] = 'X' if data.get('dust', False) else ''
            ws['D30'] = 'X' if data.get('noise', False) else ''
            ws['C31'] = 'X' if data.get('sparks', False) else ''
            ws['D31'] = data.get('other_harmful', '')
        
        # Utility Interruption
        utility = data.get('utility_interruption', 'NO')
        ws['B33'] = utility
        if utility == 'YES':
            ws['C34'] = data.get('affected_utilities', '')
    else:
        # If NO high risk, skip detailed entries
        pass
    
    # Fill Job Hazard Assessment
    jha_steps = data.get('jha_steps', [])
    if jha_steps:
        row_start = 48  # Starting row for JHA
        for i, step in enumerate(jha_steps):
            if i >= 10:  # Limit to 10 entries
                break
            ws[f'A{row_start + i}'] = step.get('step', '')
            ws[f'B{row_start + i}'] = step.get('hazard', '')
            ws[f'D{row_start + i}'] = step.get('controls', '')
    
    # Fill JHA Assessment (top section)
    ws['G3'] = data.get('jha_step1', '')
    ws['H3'] = data.get('jha_hazard1', '')
    ws['I3'] = data.get('jha_control1', '')
    ws['G5'] = data.get('jha_step2', '')
    ws['H5'] = data.get('jha_hazard2', '')
    ws['I5'] = data.get('jha_control2', '')
    
    # Fill PPE
    ppe_required = data.get('ppe_required', [])
    if 'Safety Shoes' in ppe_required:
        ws['B42'] = 'X'
    if 'Hardhat' in ppe_required:
        ws['C42'] = 'X'
    if 'Body Harness' in ppe_required:
        ws['D42'] = 'X'
    if 'Gloves' in ppe_required:
        ws['E42'] = 'X'
    if 'Welding Mask' in ppe_required:
        ws['B43'] = 'X'
    if 'N95 Masks' in ppe_required:
        ws['C43'] = 'X'
    if 'Goggles' in ppe_required:
        ws['D43'] = 'X'
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
    ws['E52'] = 'X' if data.get('approved_status') == 'YES' else ''
    ws['G52'] = data.get('safety_officer_approval', '')
    
    return wb

def get_excel_download_link(wb, filename):
    """Generate download link for Excel file"""
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    b64 = base64.b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Completed Excel File</a>'
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
        start_date = st.date_input("Start Date", value=datetime.now().date())
        end_date = st.date_input("End Date", value=datetime.now().date())
        start_time = st.time_input("Start Time", value=datetime.now().time())
        end_time = st.time_input("End Time", value=datetime.now().time())
        brief_description = st.text_area("Brief Description of Work", value="Service testing")
        
        st.subheader("🛠️ Work Details")
        
        # High Risk
        high_risk = st.radio("Is work to be done with high risk?", ["NO", "YES"])
        
        if high_risk == "YES":
            with st.expander("Work at Heights", expanded=True):
                col1a, col1b, col1c, col1d = st.columns(4)
                with col1a:
                    work_at_heights = st.checkbox("Work at Heights")
                with col1b:
                    scaffold = st.checkbox("Scaffold")
                with col1c:
                    ladder = st.checkbox("Ladder")
                with col1d:
                    tower = st.checkbox("Tower")
                
                scaffold_cert = st.text_input("NCII Cert of Scaffold Erector")
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
                scba_cert = st.text_input("Certificate of SCBA Operator")
                ventilation_eqpt = st.text_input("Ventilation Equipment")
                flash_arrester = st.checkbox("OxyFuel flash back arrester installed")
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
    
    with col2:
        st.subheader("⚠️ Job Hazard Assessment")
        
        # Main JHA Table
        st.markdown("**JOB HAZARD ASSESSMENT**")
        jha_steps = []
        
        # Step 1
        jha_step1 = st.text_input("Job Step 1", value="Site Access")
        jha_hazard1 = st.text_input("Hazard 1", value="Slip, trip, and fall from uneven surface")
        jha_control1 = st.text_area("Control 1", value="Coordinate with lessor/UDI security prior to entry. Ensure all permits are on hand. Only authorized person can enter the area")
        
        # Step 2
        jha_step2 = st.text_input("Job Step 2", value="Prepare Work Area")
        jha_hazard2 = st.text_input("Hazard 2", value="Trips/Falls: Uneven surfaces, debris, inadequate lighting")
        jha_control2 = st.text_area("Control 2", value="Clear work area of debris, ensure adequate lighting, wear appropriate footwear (safety shoes)")
        
        # Additional JHA entries
        st.markdown("---")
        st.subheader("Additional JHA Entries")
        
        num_extra = st.number_input("Number of additional JHA entries", min_value=0, max_value=10, value=3)
        
        for i in range(num_extra):
            with st.expander(f"JHA Entry {i+3}", expanded=False):
                step = st.text_input(f"Job Step {i+3}", key=f"jha_step_{i}")
                hazard = st.text_input(f"Hazard {i+3}", key=f"jha_hazard_{i}")
                controls = st.text_area(f"Controls {i+3}", key=f"jha_controls_{i}")
                if step and hazard and controls:
                    jha_steps.append({
                        'step': step,
                        'hazard': hazard,
                        'controls': controls
                    })
        
        st.subheader("🦺 Required PPE")
        ppe_options = ['Safety Shoes', 'Hardhat', 'Body Harness', 'Gloves', 'Welding Mask', 'N95 Masks', 'Goggles', 'Other PPE']
        ppe_required = st.multiselect("Select required PPE", ppe_options)
        other_ppe_text = st.text_input("Other PPE details") if 'Other PPE' in ppe_required else ""
        
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
        try:
            # Check if template exists
            if not os.path.exists('HSWP_template.xlsx'):
                st.error("⚠️ Template file 'HSWP_template.xlsx' not found in the current directory!")
                st.info("Please make sure the template file is in the same folder as this script.")
                return
            
            # Create Excel file
            wb = create_excel_template(data)
            
            # Generate download link
            filename = f"HSWP_{project_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            download_link = get_excel_download_link(wb, filename)
            
            st.success("✅ Excel file generated successfully!")
            st.markdown(download_link, unsafe_allow_html=True)
            
            # Preview data
            with st.expander("📊 Preview Data Summary"):
                st.json(data)
                
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
        - The template file 'HSWP_template.xlsx' must be in the same folder
        - All fields marked with * are required
        - The generated file will be named with the project name and date
        """)

if __name__ == "__main__":
    main()
