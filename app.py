import streamlit as st
import pandas as pd
from datetime import datetime
import io
import base64
import os
import tempfile
import zipfile
import xml.etree.ElementTree as ET

# Set page configuration
st.set_page_config(
    page_title="HSWP Automation Tool",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Health and Safety Work Permit Automation")
st.markdown("---")

def create_checkbox_overlay_xml(cell_ref, checked=False, checkbox_id=1):
    """Create XML for a checkbox overlay"""
    # This creates the XML structure for a Form Control checkbox
    x, y = cell_ref  # These would need to be calculated from cell positions
    
    checkbox_xml = f'''
    <x14:checkbox xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
        <x14:pr x14:checked="{str(checked).lower()}" x14:val="Check Box {checkbox_id}"/>
        <x14:anchor>
            <xdr:from xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
                <xdr:col>{x}</xdr:col>
                <xdr:colOff>0</xdr:colOff>
                <xdr:row>{y}</xdr:row>
                <xdr:rowOff>0</xdr:rowOff>
            </xdr:from>
            <xdr:to xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
                <xdr:col>{x+1}</xdr:col>
                <xdr:colOff>0</xdr:colOff>
                <xdr:row>{y+1}</xdr:row>
                <xdr:rowOff>0</xdr:rowOff>
            </xdr:to>
        </x14:anchor>
    </x14:checkbox>
    '''
    return checkbox_xml

def create_excel_with_overlay_checkboxes(data):
    """Create Excel file with overlay checkboxes using xlsxwriter"""
    try:
        # Use xlsxwriter for better checkbox support
        import xlsxwriter
        
        # Create a BytesIO object
        output = io.BytesIO()
        
        # Create workbook with VBA support
        workbook = xlsxwriter.Workbook(output, {'vba': True})
        
        # Add a worksheet
        worksheet = workbook.add_worksheet('Work Permit')
        
        # Set column widths
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 20)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:E', 20)
        worksheet.set_column('F:F', 20)
        worksheet.set_column('G:G', 20)
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        label_format = workbook.add_format({
            'bold': True,
            'font_size': 11
        })
        
        text_format = workbook.add_format({
            'font_size': 11
        })
        
        # Write header
        worksheet.merge_range('A1:F1', 'HEALTH and SAFETY WORK PERMIT', header_format)
        worksheet.merge_range('A2:F2', 'NOKIA SHANGHAI BELL', header_format)
        
        # PROJECT DETAILS
        worksheet.write('A4', 'Name of Sub Contractor', label_format)
        worksheet.write('B4', data.get('sub_contractor', ''), text_format)
        worksheet.write('D4', 'Requesting Vendor', label_format)
        worksheet.write('E4', data.get('requesting_vendor', ''), text_format)
        
        worksheet.write('A5', 'Project In-Charge', label_format)
        worksheet.write('B5', data.get('project_in_charge', ''), text_format)
        worksheet.write('D5', 'Person In-charge', label_format)
        worksheet.write('E5', data.get('person_in_charge', ''), text_format)
        
        worksheet.write('A6', 'Project Safety Officer', label_format)
        worksheet.write('B6', data.get('safety_officer', ''), text_format)
        worksheet.write('D6', 'Work Schedule', label_format)
        worksheet.write('E6', data.get('work_schedule', ''), text_format)
        
        worksheet.write('A7', 'Project Name', label_format)
        worksheet.write('B7', data.get('project_name', ''), text_format)
        worksheet.write('D7', 'Start Date', label_format)
        worksheet.write('E7', data.get('start_date', ''), text_format)
        worksheet.write('G7', 'Start Time', label_format)
        worksheet.write('H7', data.get('start_time', ''), text_format)
        
        worksheet.write('A8', 'Work Location', label_format)
        worksheet.write('B8', data.get('work_location', ''), text_format)
        worksheet.write('D8', 'End Date', label_format)
        worksheet.write('E8', data.get('end_date', ''), text_format)
        worksheet.write('G8', 'End Time', label_format)
        worksheet.write('H8', data.get('end_time', ''), text_format)
        
        worksheet.write('A9', 'Tower Type', label_format)
        worksheet.write('B9', data.get('tower_type', ''), text_format)
        worksheet.write('D9', 'Brief Description of Work', label_format)
        worksheet.merge_range('E9:G9', data.get('brief_description', ''), text_format)
        
        # WORK DETAILS
        worksheet.write('A11', 'WORK DETAILS', header_format)
        
        # High Risk
        worksheet.write('A12', 'Is work to be done with high risk?', label_format)
        
        # Add checkboxes for High Risk
        high_risk = data.get('high_risk', 'NO')
        # YES checkbox
        worksheet.write('C12', 'YES')
        worksheet.insert_checkbox('B12', {'checked': high_risk == 'YES'})
        # NO checkbox
        worksheet.write('D12', 'NO')
        worksheet.insert_checkbox('E12', {'checked': high_risk == 'NO'})
        
        if high_risk == 'YES':
            # Work at Heights
            worksheet.write('A13', 'Work at Heights', label_format)
            worksheet.insert_checkbox('C13', {'checked': data.get('work_at_heights', False)})
            
            worksheet.write('A14', 'Scaffold', label_format)
            worksheet.insert_checkbox('C14', {'checked': data.get('scaffold', False)})
            
            worksheet.write('A15', 'Ladder', label_format)
            worksheet.insert_checkbox('C15', {'checked': data.get('ladder', False)})
            
            worksheet.write('A16', 'Tower', label_format)
            worksheet.insert_checkbox('C16', {'checked': data.get('tower', False)})
            
            # Certifications
            worksheet.write('A17', 'NCII Cert of Scaffold Erector', label_format)
            worksheet.write('C17', data.get('scaffold_cert', ''), text_format)
            worksheet.write('D17', 'WAH Rigger Certificate', label_format)
            worksheet.write('E17', data.get('wah_rigger_cert', ''), text_format)
            
            worksheet.write('A18', 'Scaffold components available', label_format)
            worksheet.insert_checkbox('C18', {'checked': data.get('scaffold_components', False)})
            worksheet.write('D18', 'Workers physically fit', label_format)
            worksheet.insert_checkbox('E18', {'checked': data.get('workers_fit', False)})
            
            # Electrical Works
            worksheet.write('A19', 'Electrical Works', label_format)
            worksheet.insert_checkbox('C19', {'checked': data.get('electrical_works', False)})
            
            worksheet.write('A20', 'NCII Cert of Electrician or ID of REE/RME', label_format)
            worksheet.write('C20', data.get('electrician_cert', ''), text_format)
            
            worksheet.write('A21', 'LOTO Device', label_format)
            worksheet.insert_checkbox('C21', {'checked': data.get('loto_device', False)})
            worksheet.write('D21', 'Insulated Tools', label_format)
            worksheet.insert_checkbox('E21', {'checked': data.get('insulated_tools', False)})
            
            # Heavy Lifting
            worksheet.write('A22', 'Heavy Lifting w/ Equipment', label_format)
            worksheet.write('A23', 'NCII Cert of Operator', label_format)
            worksheet.write('C23', data.get('operator_cert', ''), text_format)
            worksheet.write('D23', 'Cert of Lift Rigger', label_format)
            worksheet.write('E23', data.get('rigger_cert', ''), text_format)
            worksheet.write('A24', '3rd Party Certification of Heavy Eqpt', label_format)
            worksheet.write('C24', data.get('heavy_eqpt_cert', ''), text_format)
            
            # Confined Space
            worksheet.write('A25', 'Confined Space Works', label_format)
            worksheet.insert_checkbox('C25', {'checked': data.get('confined_space', False)})
            
            worksheet.write('A26', 'Certificate of SCBA Operator', label_format)
            worksheet.write('C26', data.get('scba_cert', ''), text_format)
            worksheet.write('D26', 'Ventilation Equipment', label_format)
            worksheet.write('E26', data.get('ventilation_eqpt', ''), text_format)
            
            worksheet.write('A27', 'OxyFuel flash back arrester installed', label_format)
            worksheet.insert_checkbox('C27', {'checked': data.get('flash_arrester', False)})
            worksheet.write('D27', 'Fire Blanket', label_format)
            worksheet.insert_checkbox('E27', {'checked': data.get('fire_blanket', False)})
            
            worksheet.write('A28', 'O2 and Gas Detector', label_format)
            worksheet.write('C28', data.get('o2_detector', ''), text_format)
            worksheet.write('D28', 'Safety Line', label_format)
            worksheet.write('E28', data.get('safety_line', ''), text_format)
            
            # Harmful Substances
            worksheet.write('A29', 'Is there any harmful substance or nuisance release?', label_format)
            harmful = data.get('harmful_substance', 'NO')
            worksheet.write('C29', 'YES')
            worksheet.insert_checkbox('B29', {'checked': harmful == 'YES'})
            worksheet.write('D29', 'NO')
            worksheet.insert_checkbox('E29', {'checked': harmful == 'NO'})
            
            if harmful == 'YES':
                worksheet.write('A30', 'Fumes', label_format)
                worksheet.insert_checkbox('C30', {'checked': data.get('fumes', False)})
                worksheet.write('D30', 'Offensive Odors', label_format)
                worksheet.insert_checkbox('E30', {'checked': data.get('odors', False)})
                
                worksheet.write('A31', 'Dust', label_format)
                worksheet.insert_checkbox('C31', {'checked': data.get('dust', False)})
                worksheet.write('D31', 'Noise', label_format)
                worksheet.insert_checkbox('E31', {'checked': data.get('noise', False)})
                
                worksheet.write('A32', 'Sparks', label_format)
                worksheet.insert_checkbox('C32', {'checked': data.get('sparks', False)})
                worksheet.write('D32', 'Others:', label_format)
                worksheet.write('E32', data.get('other_harmful', ''), text_format)
            
            # Utility Interruption
            worksheet.write('A33', 'Will there be Utility interruption?', label_format)
            utility = data.get('utility_interruption', 'NO')
            worksheet.write('C33', 'YES')
            worksheet.insert_checkbox('B33', {'checked': utility == 'YES'})
            worksheet.write('D33', 'NO')
            worksheet.insert_checkbox('E33', {'checked': utility == 'NO'})
            worksheet.write('F33', 'N/A')
            worksheet.insert_checkbox('G33', {'checked': utility == 'N/A'})
            
            if utility == 'YES':
                worksheet.write('A34', 'Specify affected utilities and affected areas', label_format)
                worksheet.merge_range('C34:G34', data.get('affected_utilities', ''), text_format)
        
        # Waste Generation
        worksheet.write('A35', 'Will there be waste generation?', label_format)
        waste = data.get('waste_generation', 'NO')
        worksheet.write('C35', 'YES')
        worksheet.insert_checkbox('B35', {'checked': waste == 'YES'})
        worksheet.write('D35', 'NO')
        worksheet.insert_checkbox('E35', {'checked': waste == 'NO'})
        
        if waste == 'YES':
            worksheet.write('A36', 'Identify and list possible waste generated', label_format)
            worksheet.merge_range('C36:G36', data.get('waste_list', ''), text_format)
        
        # Tools and Materials
        worksheet.write('A38', 'List of tools and materials to be used', label_format)
        tools = data.get('tools_materials', [])
        for i, tool in enumerate(tools[:10]):
            worksheet.write(f'A{39+i}', tool, text_format)
        
        # PPE
        worksheet.write('A45', 'REQUIRED PPE', header_format)
        ppe_row = 46
        ppe_items = ['Safety Shoes', 'Hardhat', 'Body Harness', 'Gloves', 'Welding Mask', 'N95 Masks', 'Goggles', 'Other PPE']
        ppe_selected = data.get('ppe_required', [])
        
        for i, ppe in enumerate(ppe_items):
            col = chr(65 + i)  # A, B, C, D, etc.
            worksheet.write(f'{col}{ppe_row}', ppe, label_format)
            if ppe in ppe_selected:
                worksheet.insert_checkbox(f'{col}{ppe_row+1}', {'checked': True})
            else:
                worksheet.insert_checkbox(f'{col}{ppe_row+1}', {'checked': False})
        
        if 'Other PPE' in ppe_selected:
            worksheet.write('E47', data.get('other_ppe_text', ''), text_format)
        
        # Workers
        worksheet.write('A50', 'List of workers', label_format)
        workers = data.get('workers', [])
        for i, worker in enumerate(workers[:8]):
            worksheet.write(f'A{51+i}', worker, text_format)
        
        # Acknowledgement
        worksheet.write('A58', 'ACKNOWLEDGEMENT', header_format)
        worksheet.write('A59', 'Prepared By:', label_format)
        worksheet.write('B59', data.get('prepared_by', ''), text_format)
        worksheet.write('C59', 'Noted By', label_format)
        worksheet.write('D59', data.get('noted_by', ''), text_format)
        worksheet.write('E59', 'Approved By', label_format)
        
        # Approval checkbox
        worksheet.insert_checkbox('G59', {'checked': data.get('approved_status') == 'YES'})
        
        worksheet.write('A60', 'Project Safety Officer', label_format)
        worksheet.write('B60', 'MNO/Project In-Charge', label_format)
        worksheet.write('D60', data.get('safety_officer_approval', ''), text_format)
        
        # Close workbook
        workbook.close()
        
        # Get the file data
        output.seek(0)
        file_data = output.getvalue()
        
        return io.BytesIO(file_data)
        
    except Exception as e:
        st.error(f"Error creating Excel with checkboxes: {str(e)}")
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
        
        # Main JHA Table
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
    
    # Important notice about checkboxes
    st.info("""
    ℹ️ **About Checkboxes:** 
    - This creates **interactive overlay checkboxes** (Form Controls)
    - These float above the cells and can be clicked
    - **Note:** You may need to enable editing/design mode in Excel to fully interact with them
    - File is created using xlsxwriter for better checkbox support
    """)
    
    # Generate Excel button
    if st.button("📥 Generate Excel File", type="primary"):
        with st.spinner("Generating Excel file with overlay checkboxes..."):
            try:
                # First try with xlsxwriter
                try:
                    import xlsxwriter
                    file_data = create_excel_with_overlay_checkboxes(data)
                except ImportError:
                    st.warning("xlsxwriter not installed. Falling back to openpyxl with Unicode checkboxes.")
                    # Fallback to openpyxl with Unicode checkboxes
                    file_data = create_excel_template_with_unicode(data)
                
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
        5. Download the completed file
        """)
        
        st.header("📋 Requirements")
        st.markdown("""
        - Python package: `xlsxwriter`
        - Install with: `pip install xlsxwriter`
        """)
        
        st.header("ℹ️ About Checkboxes")
        st.markdown("""
        This creates **overlay checkboxes** that:
        - Float above the cells
        - Are interactive (clickable)
        - Are Excel Form Controls
        - Work like the ones in your template
        """)

if __name__ == "__main__":
    main()
