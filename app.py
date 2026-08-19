import streamlit as st
import pandas as pd
from datetime import datetime
import io
import base64
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
import os
import shutil
import tempfile
import subprocess
import sys

# Set page configuration
st.set_page_config(
    page_title="HSWP Automation Tool",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Health and Safety Work Permit Automation")
st.markdown("---")

def get_top_left_cell(ws, cell_address):
    """Get the top-left cell of a merged range if the cell is merged"""
    try:
        cell = ws[cell_address]
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
                return ws.cell(row=min_row, column=min_col)
        return cell
    except:
        return ws[cell_address]

def safe_write_cell(ws, cell_address, value):
    """Safely write a value to a cell, handling merged cells"""
    try:
        target_cell = get_top_left_cell(ws, cell_address)
        if target_cell:
            target_cell.value = value
        return True
    except:
        return False

def create_excel_template(data):
    """Create Excel file with data and VBA macro to add checkboxes"""
    
    # Check if template exists
    template_path = 'HSWP_template.xlsx'
    if not os.path.exists(template_path):
        st.error("⚠️ Template file 'HSWP_template.xlsx' not found!")
        return None
    
    try:
        # Create a temporary copy of the template
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            shutil.copy2(template_path, tmp_file.name)
            temp_template_path = tmp_file.name
        
        # Load the template
        wb = load_workbook(temp_template_path)
        ws = wb.worksheets[0]
        
        # Fill PROJECT DETAILS
        safe_write_cell(ws, 'B2', data.get('sub_contractor', ''))
        safe_write_cell(ws, 'D2', data.get('requesting_vendor', ''))
        safe_write_cell(ws, 'B4', data.get('project_in_charge', ''))
        safe_write_cell(ws, 'B5', data.get('safety_officer', ''))
        safe_write_cell(ws, 'B6', data.get('project_name', ''))
        safe_write_cell(ws, 'B7', data.get('work_location', ''))
        safe_write_cell(ws, 'B8', data.get('tower_type', ''))
        safe_write_cell(ws, 'D4', data.get('person_in_charge', ''))
        safe_write_cell(ws, 'D5', data.get('work_schedule', ''))
        safe_write_cell(ws, 'D6', data.get('start_date', ''))
        safe_write_cell(ws, 'D7', data.get('end_date', ''))
        safe_write_cell(ws, 'F5', data.get('work_time_period', ''))
        safe_write_cell(ws, 'F6', data.get('start_time', ''))
        safe_write_cell(ws, 'F7', data.get('end_time', ''))
        safe_write_cell(ws, 'E8', data.get('brief_description', ''))
        
        # Fill JHA Assessment
        safe_write_cell(ws, 'G3', data.get('jha_step1', ''))
        safe_write_cell(ws, 'H3', data.get('jha_hazard1', ''))
        safe_write_cell(ws, 'I3', data.get('jha_control1', ''))
        safe_write_cell(ws, 'G5', data.get('jha_step2', ''))
        safe_write_cell(ws, 'H5', data.get('jha_hazard2', ''))
        safe_write_cell(ws, 'I5', data.get('jha_control2', ''))
        
        # Fill checkbox-linked cells with X or empty
        # High Risk
        high_risk = data.get('high_risk', 'NO')
        safe_write_cell(ws, 'B12', 'X' if high_risk == 'YES' else '')
        safe_write_cell(ws, 'D12', 'X' if high_risk == 'NO' else '')
        
        # Work at Heights
        safe_write_cell(ws, 'C13', 'X' if data.get('work_at_heights', False) else '')
        safe_write_cell(ws, 'C14', 'X' if data.get('scaffold', False) else '')
        safe_write_cell(ws, 'C15', 'X' if data.get('ladder', False) else '')
        safe_write_cell(ws, 'C16', 'X' if data.get('tower', False) else '')
        
        # Certifications
        safe_write_cell(ws, 'C13', data.get('scaffold_cert', ''))
        safe_write_cell(ws, 'E13', data.get('wah_rigger_cert', ''))
        safe_write_cell(ws, 'C18', 'X' if data.get('scaffold_components', False) else '')
        safe_write_cell(ws, 'E18', 'X' if data.get('workers_fit', False) else '')
        
        # Electrical Works
        safe_write_cell(ws, 'C19', 'X' if data.get('electrical_works', False) else '')
        safe_write_cell(ws, 'C20', data.get('electrician_cert', ''))
        safe_write_cell(ws, 'C21', 'X' if data.get('loto_device', False) else '')
        safe_write_cell(ws, 'E21', 'X' if data.get('insulated_tools', False) else '')
        
        # Heavy Lifting
        safe_write_cell(ws, 'C23', data.get('operator_cert', ''))
        safe_write_cell(ws, 'E23', data.get('rigger_cert', ''))
        safe_write_cell(ws, 'C24', data.get('heavy_eqpt_cert', ''))
        
        # Confined Space
        safe_write_cell(ws, 'C25', 'X' if data.get('confined_space', False) else '')
        safe_write_cell(ws, 'C26', data.get('scba_cert', ''))
        safe_write_cell(ws, 'E26', data.get('ventilation_eqpt', ''))
        safe_write_cell(ws, 'C27', 'X' if data.get('flash_arrester', False) else '')
        safe_write_cell(ws, 'E27', 'X' if data.get('fire_blanket', False) else '')
        safe_write_cell(ws, 'C28', data.get('o2_detector', ''))
        safe_write_cell(ws, 'E28', data.get('safety_line', ''))
        
        # Harmful Substances
        harmful = data.get('harmful_substance', 'NO')
        safe_write_cell(ws, 'B29', 'X' if harmful == 'YES' else '')
        safe_write_cell(ws, 'D29', 'X' if harmful == 'NO' else '')
        safe_write_cell(ws, 'C30', 'X' if data.get('fumes', False) else '')
        safe_write_cell(ws, 'E30', 'X' if data.get('odors', False) else '')
        safe_write_cell(ws, 'C31', 'X' if data.get('dust', False) else '')
        safe_write_cell(ws, 'E31', 'X' if data.get('noise', False) else '')
        safe_write_cell(ws, 'C32', 'X' if data.get('sparks', False) else '')
        safe_write_cell(ws, 'E32', data.get('other_harmful', ''))
        
        # Utility Interruption
        utility = data.get('utility_interruption', 'NO')
        safe_write_cell(ws, 'B33', 'X' if utility == 'YES' else '')
        safe_write_cell(ws, 'D33', 'X' if utility == 'NO' else '')
        safe_write_cell(ws, 'F33', 'X' if utility == 'N/A' else '')
        safe_write_cell(ws, 'C34', data.get('affected_utilities', ''))
        
        # Waste Generation
        waste = data.get('waste_generation', 'NO')
        safe_write_cell(ws, 'B35', 'X' if waste == 'YES' else '')
        safe_write_cell(ws, 'D35', 'X' if waste == 'NO' else '')
        safe_write_cell(ws, 'C36', data.get('waste_list', ''))
        
        # PPE - mark with X
        ppe_required = data.get('ppe_required', [])
        ppe_cells = {
            'Safety Shoes': 'B46',
            'Hardhat': 'C46',
            'Body Harness': 'D46',
            'Gloves': 'E46',
            'Welding Mask': 'B47',
            'N95 Masks': 'C47',
            'Goggles': 'D47'
        }
        for ppe, cell in ppe_cells.items():
            safe_write_cell(ws, cell, 'X' if ppe in ppe_required else '')
        
        if 'Other PPE' in ppe_required:
            safe_write_cell(ws, 'E47', data.get('other_ppe_text', ''))
        
        # Fill JHA Table
        jha_steps = data.get('jha_steps', [])
        row_start = 48
        for i, step in enumerate(jha_steps):
            if i >= 10:
                break
            safe_write_cell(ws, f'A{row_start + i}', step.get('step', ''))
            safe_write_cell(ws, f'B{row_start + i}', step.get('hazard', ''))
            safe_write_cell(ws, f'D{row_start + i}', step.get('controls', ''))
        
        # Fill Tools and Materials
        tools = data.get('tools_materials', [])
        tool_row_start = 39
        for i, tool in enumerate(tools):
            if i >= 10:
                break
            safe_write_cell(ws, f'A{tool_row_start + i}', tool)
        
        # Fill Workers
        workers = data.get('workers', [])
        worker_row_start = 44
        for i, worker in enumerate(workers):
            if i >= 8:
                break
            safe_write_cell(ws, f'A{worker_row_start + i}', worker)
        
        # Fill Acknowledgement
        safe_write_cell(ws, 'B51', data.get('prepared_by', ''))
        safe_write_cell(ws, 'C51', data.get('noted_by', ''))
        safe_write_cell(ws, 'E51', data.get('approved_by', ''))
        safe_write_cell(ws, 'C52', data.get('noted_by', ''))
        safe_write_cell(ws, 'F59', 'X' if data.get('approved_status') == 'YES' else '')
        safe_write_cell(ws, 'G52', data.get('safety_officer_approval', ''))
        
        # Save the workbook
        wb.save(temp_template_path)
        
        # Read the saved file into memory for download
        with open(temp_template_path, 'rb') as f:
            file_data = f.read()
        
        # Clean up temp file
        try:
            os.unlink(temp_template_path)
        except:
            pass
        
        output = io.BytesIO(file_data)
        output.seek(0)
        
        return output
        
    except Exception as e:
        st.error(f"Error processing template: {str(e)}")
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
        
        brief_description = st.text_area("Brief Description of Work", value="Service testing", height=68)
        
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
    
    with col2:
        st.subheader("⚠️ Job Hazard Assessment")
        
        st.markdown("**JOB HAZARD ASSESSMENT**")
        
        jha_step1 = st.text_input("Job Step 1", value="Site Access")
        jha_hazard1 = st.text_input("Hazard 1", value="Slip, trip, and fall from uneven surface")
        jha_control1 = st.text_area("Control 1", value="Coordinate with lessor/UDI security prior to entry. Ensure all permits are on hand. Only authorized person can enter the area", height=60)
        
        jha_step2 = st.text_input("Job Step 2", value="Prepare Work Area")
        jha_hazard2 = st.text_input("Hazard 2", value="Trips/Falls: Uneven surfaces, debris, inadequate lighting")
        jha_control2 = st.text_area("Control 2", value="Clear work area of debris, ensure adequate lighting, wear appropriate footwear (safety shoes)", height=60)
        
        jha_step3 = st.text_input("Job Step 3", value="Handling of Fiber Optic cables")
        jha_hazard3 = st.text_input("Hazard 3", value="Cuts/Lacerations: Sharp edges of fibers or connectors")
        jha_control3 = st.text_area("Control 3", value="Wear cut-resistant gloves. Handle fibers with care", height=60)
        
        st.markdown("---")
        st.subheader("Additional JHA Entries")
        
        jha_steps = []
        
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
                                  value="PPE\nFIRST AID KIT\nODF\nOPM\nOTDR\nPatchcord\nFusion Splicer\nCleaning Kit", height=120)
        tools = [t.strip() for t in tools_text.split('\n') if t.strip()]
        
        st.subheader("👷 List of Workers")
        workers_text = st.text_area("List workers (one per line)",
                                   value="JAY PALASOL\nROSE EISELE BARBA\nEDGAR PERALTA\nMELVIN ADOVE\nROCKY MARZO TRIVIÑO\nJANN ALEXIS AGULO\nRODERICK REYES\nMARK JOSEPH INOVERO\nMARK JAYSON BRILLANTES", height=160)
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
        'waste_generation': waste_generation,
        'waste_list': waste_list,
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
    
    st.info("""
    ℹ️ **How Checkboxes Work:**
    1. The template contains a VBA macro that creates overlay checkboxes
    2. Data is filled using openpyxl (preserves all formatting)
    3. The macro reads values from linked cells (X = checked, empty = unchecked)
    4. **To activate checkboxes:** Open the file in Excel, enable macros, and run the `AddCheckboxes` macro
    5. Or use the macro-enabled template version
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
                    
                    st.info("""
                    📌 **Next Steps:**
                    1. Download and open the file in Excel
                    2. **Enable Macros** when prompted
                    3. Press `Alt+F8` or go to Developer > Macros
                    4. Select `AddCheckboxes` and click Run
                    5. The overlay checkboxes will appear!
                    """)
                    
                    with st.expander("📊 Preview Data Summary"):
                        summary_data = {
                            'Project': project_name,
                            'Location': work_location,
                            'Sub Contractor': sub_contractor,
                            'Safety Officer': safety_officer,
                            'High Risk': high_risk,
                            'Workers': len(workers),
                            'JHA Steps': len(jha_steps),
                            'Tools': len(tools),
                            'PPE Items': len(ppe_required)
                        }
                        st.json(summary_data)
                else:
                    st.error("❌ Failed to generate Excel file.")
                    
            except Exception as e:
                st.error(f"❌ Error generating file: {str(e)}")
                st.exception(e)
    
    # Instructions
    with st.sidebar:
        st.header("📝 Instructions")
        st.markdown("""
        1. Fill in all required fields
        2. Add JHA entries as needed
        3. Select required PPE
        4. Click 'Generate Excel File'
        5. Download and open in Excel
        6. Enable macros
        7. Run the AddCheckboxes macro
        """)
        
        st.header("📋 Template Requirements")
        st.markdown("""
        - Template file: `HSWP_template.xlsx`
        - Must be in the repository root
        - **Enable Macros** in Excel
        - The original template is preserved
        """)
        
        st.header("🔧 How It Works")
        st.markdown("""
        1. **openpyxl** fills data in the template
        2. **Linked cells** store checkbox states (X = checked)
        3. **VBA macro** reads the linked cells and creates checkboxes
        4. **Overlay checkboxes** appear on top of the cells
        """)

if __name__ == "__main__":
    main()
