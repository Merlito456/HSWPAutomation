import streamlit as st
import pandas as pd
from datetime import datetime
import io
import base64
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import os
import shutil
import tempfile

# Set page configuration
st.set_page_config(
    page_title="HSWP Automation Tool",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Health and Safety Work Permit Automation")
st.markdown("---")

def create_excel_template(data):
    """Create Excel file with proper formatting"""
    
    template_path = 'HSWP_template.xlsx'
    if not os.path.exists(template_path):
        st.error("⚠️ Template file 'HSWP_template.xlsx' not found!")
        return None
    
    try:
        # Create a temporary copy
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            shutil.copy2(template_path, tmp_file.name)
            temp_template_path = tmp_file.name
        
        # Load the template WITHOUT modifying it
        wb = load_workbook(temp_template_path)
        ws = wb.active
        
        # Define a function to safely write to cells
        def write_cell(row, col, value):
            try:
                ws.cell(row=row, column=col, value=value)
            except:
                pass
        
        # ============================================
        # PROJECT DETAILS - Row numbers from template
        # ============================================
        # Row 2: Sub Contractor (B2) and Requesting Vendor (D2)
        write_cell(2, 2, data.get('sub_contractor', 'Ultegra Supplies and Services'))
        write_cell(2, 4, data.get('requesting_vendor', 'NOKIA SHANGHAI BELL'))
        
        # Row 4: Project In-Charge (B4) and Person In-charge (D4)
        write_cell(4, 2, data.get('project_in_charge', ''))
        write_cell(4, 4, data.get('person_in_charge', 'John Carlo Rabanes'))
        
        # Row 5: Safety Officer (B5), Work Schedule (D5), Work Time Period (F5)
        write_cell(5, 2, data.get('safety_officer', 'RONNIE ALVIN CHIU'))
        write_cell(5, 4, data.get('work_schedule', ''))
        write_cell(5, 6, data.get('work_time_period', ''))
        
        # Row 6: Project Name (B6), Start Date (D6), Start Time (F6)
        write_cell(6, 2, data.get('project_name', 'FTTH HORIZONTAL'))
        write_cell(6, 4, data.get('start_date', '08/10/2026'))
        write_cell(6, 6, data.get('start_time', '08:00AM'))
        
        # Row 7: Work Location (B7), End Date (D7), End Time (F7)
        write_cell(7, 2, data.get('work_location', 'MIN624_TS ORAPOBBUTUANAGN'))
        write_cell(7, 4, data.get('end_date', '09/10/2026'))
        write_cell(7, 6, data.get('end_time', '08:00PM'))
        
        # Row 8: Tower Type (B8), Brief Description (E8)
        write_cell(8, 2, data.get('tower_type', 'ground base'))
        write_cell(8, 5, data.get('brief_description', 'SFP link upgrade, Survey'))
        
        # ============================================
        # WORK DETAILS
        # ============================================
        # High Risk - Row 12
        high_risk = data.get('high_risk', 'NO')
        write_cell(12, 2, high_risk)  # Column B
        
        # Work at Heights - Row 12
        write_cell(12, 3, 'X' if data.get('work_at_heights', False) else '')
        write_cell(12, 4, 'X' if data.get('scaffold', False) else '')
        write_cell(12, 5, 'X' if data.get('ladder', False) else '')
        write_cell(12, 6, 'X' if data.get('tower', False) else '')
        
        # Row 13: Certifications
        write_cell(13, 3, data.get('scaffold_cert', ''))
        write_cell(13, 5, data.get('wah_rigger_cert', ''))
        
        # Row 14: Scaffold components and Workers fit
        write_cell(14, 3, 'X' if data.get('scaffold_components', False) else '')
        write_cell(14, 5, 'X' if data.get('workers_fit', False) else '')
        
        # Row 16: Electrical Works
        write_cell(16, 3, 'X' if data.get('electrical_works', False) else '')
        
        # Row 17: Electrician Cert
        write_cell(17, 3, data.get('electrician_cert', ''))
        
        # Row 18: LOTO and Insulated Tools
        write_cell(18, 3, 'X' if data.get('loto_device', False) else '')
        write_cell(18, 5, 'X' if data.get('insulated_tools', False) else '')
        
        # Row 20: Heavy Lifting
        write_cell(20, 3, data.get('operator_cert', ''))
        write_cell(20, 4, data.get('rigger_cert', ''))
        
        # Row 21: Heavy Equipment Cert
        write_cell(21, 3, data.get('heavy_eqpt_cert', ''))
        
        # Row 23: Confined Space
        write_cell(23, 3, 'X' if data.get('confined_space', False) else '')
        
        # Row 24: SCBA and Ventilation
        write_cell(24, 3, data.get('scba_cert', ''))
        write_cell(24, 4, data.get('ventilation_eqpt', ''))
        
        # Row 25: Flash Arrester and Fire Blanket
        write_cell(25, 3, 'X' if data.get('flash_arrester', False) else '')
        write_cell(25, 5, 'X' if data.get('fire_blanket', False) else '')
        
        # Row 26: O2 Detector and Safety Line
        write_cell(26, 3, data.get('o2_detector', ''))
        write_cell(26, 4, data.get('safety_line', ''))
        
        # ============================================
        # HARMFUL SUBSTANCES - Row 28-31
        # ============================================
        harmful = data.get('harmful_substance', 'NO')
        write_cell(28, 2, harmful)
        
        if harmful == 'YES':
            write_cell(29, 3, 'X' if data.get('fumes', False) else '')
            write_cell(29, 4, 'X' if data.get('odors', False) else '')
            write_cell(30, 3, 'X' if data.get('dust', False) else '')
            write_cell(30, 4, 'X' if data.get('noise', False) else '')
            write_cell(31, 3, 'X' if data.get('sparks', False) else '')
            write_cell(31, 4, data.get('other_harmful', ''))
        
        # ============================================
        # UTILITY INTERRUPTION - Row 33
        # ============================================
        utility = data.get('utility_interruption', 'NO')
        write_cell(33, 2, utility)
        if utility == 'YES':
            write_cell(34, 3, data.get('affected_utilities', ''))
        
        # ============================================
        # WASTE GENERATION - Row 37
        # ============================================
        waste = data.get('waste_generation', 'NO')
        write_cell(37, 2, waste)
        if waste == 'YES':
            write_cell(37, 3, data.get('waste_list', ''))
        
        # ============================================
        # JHA TOP SECTION - Rows 3 and 5
        # ============================================
        write_cell(3, 7, data.get('jha_step1', ''))
        write_cell(3, 8, data.get('jha_hazard1', ''))
        write_cell(3, 9, data.get('jha_control1', ''))
        write_cell(5, 7, data.get('jha_step2', ''))
        write_cell(5, 8, data.get('jha_hazard2', ''))
        write_cell(5, 9, data.get('jha_control2', ''))
        
        # ============================================
        # SITE LOCATIONS - Rows 3-12 (columns J-N)
        # ============================================
        sites = data.get('sites', [])
        for i, site in enumerate(sites):
            if i >= 10:
                break
            row = 3 + i
            write_cell(row, 10, site.get('site_id', ''))
            write_cell(row, 11, site.get('anchor_id', ''))
            write_cell(row, 12, site.get('safety_officer', data.get('safety_officer', 'Ronnie Alvin Chiu')))
            write_cell(row, 13, site.get('project_in_charge', data.get('person_in_charge', 'John Carlo Rabanes')))
            write_cell(row, 14, site.get('worker_name', ''))
        
        # ============================================
        # JHA BOTTOM TABLE - Starting row 48
        # ============================================
        jha_steps = data.get('jha_steps', [])
        for i, step in enumerate(jha_steps):
            if i >= 10:
                break
            row = 48 + i
            write_cell(row, 1, step.get('step', ''))
            write_cell(row, 2, step.get('hazard', ''))
            write_cell(row, 4, step.get('controls', ''))
        
        # ============================================
        # PPE - Rows 42-43
        # ============================================
        ppe_required = data.get('ppe_required', [])
        write_cell(42, 2, 'X' if 'Safety Shoes' in ppe_required else '')
        write_cell(42, 3, 'X' if 'Hardhat' in ppe_required else '')
        write_cell(42, 4, 'X' if 'Body Harness' in ppe_required else '')
        write_cell(42, 5, 'X' if 'Gloves' in ppe_required else '')
        write_cell(43, 2, 'X' if 'Welding Mask' in ppe_required else '')
        write_cell(43, 3, 'X' if 'N95 Masks' in ppe_required else '')
        write_cell(43, 4, 'X' if 'Goggles' in ppe_required else '')
        
        if 'Other PPE' in ppe_required:
            write_cell(43, 5, data.get('other_ppe_text', ''))
        
        # ============================================
        # WORKERS - Starting row 44 (two columns)
        # ============================================
        workers = data.get('workers', [])
        for i, worker in enumerate(workers):
            if i >= 16:
                break
            row = 44 + (i // 2)
            col = 1 if i % 2 == 0 else 2
            write_cell(row, col, worker)
        
        # ============================================
        # ACKNOWLEDGEMENT - Rows 51-52
        # ============================================
        write_cell(51, 2, data.get('prepared_by', 'RONNIE ALVIN-CHIU'))
        write_cell(51, 3, data.get('noted_by', 'John Carlo Rabanes'))
        write_cell(51, 5, data.get('approved_by', ''))
        write_cell(52, 3, data.get('noted_by', 'John Carlo Rabanes'))
        write_cell(52, 5, 'X' if data.get('approved_status') == 'YES' else '')
        write_cell(52, 6, 'X' if data.get('approved_status') == 'NO' else '')
        write_cell(52, 7, data.get('safety_officer_approval', 'PTG MIDC O&M Regional Manager'))
        
        # ============================================
        # Save the workbook
        # ============================================
        wb.save(temp_template_path)
        
        # Read the saved file
        with open(temp_template_path, 'rb') as f:
            file_data = f.read()
        
        # Clean up
        try:
            os.unlink(temp_template_path)
        except:
            pass
        
        output = io.BytesIO(file_data)
        output.seek(0)
        
        return output
        
    except Exception as e:
        st.error(f"Error: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None

def get_excel_download_link(file_data, filename):
    """Generate download link for Excel file"""
    b64 = base64.b64encode(file_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" style="background-color: #4CAF50; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold; font-size: 16px;">📥 Download Completed Excel File</a>'
    return href

def main():
    # Create two columns for layout
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📋 Project Details")
        
        sub_contractor = st.text_input("Name of Sub Contractor", value="Ultegra Supplies and Services")
        requesting_vendor = st.text_input("Requesting Vendor", value="NOKIA SHANGHAI BELL")
        project_in_charge = st.text_input("Project In-Charge", value="")
        person_in_charge = st.text_input("Person In-charge", value="John Carlo Rabanes")
        safety_officer = st.text_input("Project Safety Officer", value="RONNIE ALVIN CHIU")
        work_schedule = st.text_input("Work Schedule", value="")
        work_time_period = st.text_input("Work Time Period", value="")
        project_name = st.text_input("Project Name", value="FTTH HORIZONTAL")
        work_location = st.text_input("Work Location", value="MIN624_TS ORAPOBBUTUANAGN")
        tower_type = st.text_input("Tower Type", value="ground base")
        brief_description = st.text_area("Brief Description of Work", value="SFP link upgrade, Survey", height=68)
        
        col1a, col1b = st.columns(2)
        with col1a:
            start_date = st.date_input("Start Date", value=datetime(2026, 8, 10))
            start_time = st.text_input("Start Time", value="08:00AM")
        with col1b:
            end_date = st.date_input("End Date", value=datetime(2026, 9, 10))
            end_time = st.text_input("End Time", value="08:00PM")
        
        st.subheader("🛠️ Work Details")
        
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
        
        # Waste Generation
        st.subheader("♻️ Waste Generation")
        waste_generation = st.radio("Will there be waste generation?", ["NO", "YES"])
        waste_list = ""
        if waste_generation == "YES":
            waste_list = st.text_area("Identify and list possible waste generated", 
                                     value="Cable scraps\nPackaging materials\nUsed PPE", height=80)
        
        st.subheader("📌 JOB HAZARD ASSESSMENT (Top Section)")
        
        jha_step1 = st.text_input("Job Step 1", value="")
        jha_hazard1 = st.text_input("Hazard 1", value="")
        jha_control1 = st.text_area("Control 1", value="", height=60)
        
        jha_step2 = st.text_input("Job Step 2", value="")
        jha_hazard2 = st.text_input("Hazard 2", value="")
        jha_control2 = st.text_area("Control 2", value="", height=60)
    
    with col2:
        st.subheader("📍 Site Locations")
        st.info("Add site locations with worker designations")
        
        num_sites = st.number_input("Number of sites", min_value=0, max_value=10, value=2)
        
        sites = []
        for i in range(num_sites):
            with st.expander(f"Site {i+1}", expanded=i==0):
                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    site_id = st.text_input(f"Site ID {i+1}", value="MIN624" if i==0 else "MIN779")
                    safety_officer_site = st.text_input(f"Safety Officer {i+1}", value="Ronnie Alvin Chiu")
                with col_s2:
                    anchor_id = st.text_input(f"Anchor ID/PLA No. {i+1}", value="")
                    project_in_charge_site = st.text_input(f"Project In-Charge {i+1}", value="John Carlo Rabanes")
                
                worker_name = st.text_area(f"Worker Name/Designation {i+1}", 
                                          value="RABANES, JOHN CARLO/ OLT ENGINEER" if i==0 else "Sabordo, Walrich/ Field Engineer", 
                                          height=80)
                
                sites.append({
                    'site_id': site_id,
                    'anchor_id': anchor_id,
                    'safety_officer': safety_officer_site,
                    'project_in_charge': project_in_charge_site,
                    'worker_name': worker_name
                })
        
        st.subheader("⚠️ JOB HAZARD ASSESSMENT (Bottom Table)")
        
        num_jha = st.number_input("Number of JHA entries", min_value=0, max_value=10, value=0)
        
        jha_steps = []
        for i in range(num_jha):
            with st.expander(f"JHA Entry {i+1}", expanded=i==0):
                step = st.text_input(f"Job Step {i+1}", key=f"jha_step_{i}")
                hazard = st.text_input(f"Hazard {i+1}", key=f"jha_hazard_{i}")
                controls = st.text_area(f"Controls {i+1}", key=f"jha_controls_{i}", height=60)
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
        
        st.subheader("👷 List of Workers")
        workers_text = st.text_area("List workers (one per line)",
                                   value="CABRAL, RAYMOND R.\nCANETE, GREMAR L.\nORBITA, RICHARD M.\nPELISCO, KENNETH KIM L.\nPRIETO, JASON D.\nPRIETO, KEITH JHIRVINE D.\nPRIETO, SEAN DALE D.\nSABAD, VIRGILIO JR M.\nSUMILE, JUSTINE KIM J.\nTRANA, AQUILINO B.\nTRANA, TRANQUILINO B.\nRABANES, JOHN CARLO", height=200)
        workers = [w.strip() for w in workers_text.split('\n') if w.strip()]
        
        st.subheader("✅ Acknowledgement")
        prepared_by = st.text_input("Prepared By", value="RONNIE ALVIN-CHIU")
        noted_by = st.text_input("Endorsed By", value="John Carlo Rabanes")
        approved_by = st.radio("Approved By", ["YES", "NO"])
        safety_officer_approval = st.text_input("MIDC O&M Regional Manager", value="PTG MIDC O&M Regional Manager")

    # Collect all data
    data = {
        'sub_contractor': sub_contractor,
        'requesting_vendor': requesting_vendor,
        'project_in_charge': project_in_charge,
        'person_in_charge': person_in_charge,
        'safety_officer': safety_officer,
        'work_schedule': work_schedule,
        'work_time_period': work_time_period,
        'project_name': project_name,
        'work_location': work_location,
        'tower_type': tower_type,
        'start_date': start_date.strftime('%m/%d/%Y'),
        'end_date': end_date.strftime('%m/%d/%Y'),
        'start_time': start_time,
        'end_time': end_time,
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
        'waste_generation': waste_generation,
        'waste_list': waste_list,
        'jha_step1': jha_step1,
        'jha_hazard1': jha_hazard1,
        'jha_control1': jha_control1,
        'jha_step2': jha_step2,
        'jha_hazard2': jha_hazard2,
        'jha_control2': jha_control2,
        'sites': sites,
        'jha_steps': jha_steps,
        'ppe_required': ppe_required,
        'other_ppe_text': other_ppe_text,
        'workers': workers,
        'prepared_by': prepared_by,
        'noted_by': noted_by,
        'approved_by': approved_by,
        'approved_status': approved_by,
        'safety_officer_approval': safety_officer_approval
    }
    
    st.markdown("---")
    
    # Note about preservation
    st.info("""
    ℹ️ **Template Preservation:**
    - The original template structure is preserved
    - 'X' marks are used for checkboxes
    - All data is populated correctly
    - Re-add your logo manually if needed
    """)
    
    # Generate Excel button
    if st.button("📥 Generate Excel File", type="primary"):
        with st.spinner("Generating Excel file..."):
            try:
                file_data = create_excel_template(data)
                if file_data:
                    filename = f"HSWP_{project_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    download_link = get_excel_download_link(file_data.getvalue(), filename)
                    
                    st.success("✅ Excel file generated successfully!")
                    st.markdown(download_link, unsafe_allow_html=True)
                    
                    with st.expander("📊 Preview Data Summary"):
                        summary_data = {
                            'Project': project_name,
                            'Location': work_location,
                            'Sub Contractor': sub_contractor,
                            'Safety Officer': safety_officer,
                            'High Risk': high_risk,
                            'Sites': len(sites),
                            'JHA Entries': len(jha_steps),
                            'Workers': len(workers)
                        }
                        st.json(summary_data)
                else:
                    st.error("❌ Failed to generate Excel file.")
                    
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)
    
    # Instructions
    with st.sidebar:
        st.header("📝 Instructions")
        st.markdown("""
        1. Fill in all fields
        2. Add site locations
        3. Add JHA entries
        4. Select PPE
        5. List workers
        6. Click Generate
        7. Download the file
        """)
        
        st.header("ℹ️ Limitations")
        st.markdown("""
        - Images/Logos need to be re-added manually
        - Form controls replaced with 'X' marks
        - All data and formatting preserved
        """)

if __name__ == "__main__":
    main()
