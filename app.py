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
    """Create Excel file with filled data - preserving original template"""
    
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
        ws = wb.worksheets[0]  # First sheet (Work Permit)
        
        # ============================================
        # PROJECT DETAILS - Only modify these fields
        # ============================================
        
        # Row 2: Sub Contractor and Requesting Vendor
        safe_write_cell(ws, 'B2', data.get('sub_contractor', 'DNA'))
        safe_write_cell(ws, 'D2', data.get('requesting_vendor', 'NOKIA'))
        
        # Row 4: Project In-Charge and Person In-charge
        safe_write_cell(ws, 'B4', data.get('project_in_charge', ''))
        safe_write_cell(ws, 'D4', data.get('person_in_charge', 'John Carlo Rabanes'))
        
        # Row 5: Safety Officer and Work Schedule/Time Period
        safe_write_cell(ws, 'B5', data.get('safety_officer', 'RONNIE ALVIN CHIU'))
        safe_write_cell(ws, 'D5', data.get('work_schedule', ''))
        safe_write_cell(ws, 'F5', data.get('work_time_period', ''))
        
        # Row 6: Project Name, Start Date, Start Time
        safe_write_cell(ws, 'B6', data.get('project_name', 'FTTH HORIZONTAL'))
        safe_write_cell(ws, 'D6', data.get('start_date', ''))
        safe_write_cell(ws, 'F6', data.get('start_time', ''))
        
        # Row 7: Work Location, End Date, End Time
        safe_write_cell(ws, 'B7', data.get('work_location', 'MIN624_TS ORAPOBBUTUANAGN'))
        safe_write_cell(ws, 'D7', data.get('end_date', ''))
        safe_write_cell(ws, 'F7', data.get('end_time', ''))
        
        # Row 8: Tower Type and Brief Description
        safe_write_cell(ws, 'B8', data.get('tower_type', 'ground base'))
        safe_write_cell(ws, 'E8', data.get('brief_description', 'SFP link upgrade, Survey'))
        
        # ============================================
        # JOB HAZARD ASSESSMENT (Top section)
        # ============================================
        # Row 3: JHA Step 1
        safe_write_cell(ws, 'G3', data.get('jha_step1', ''))
        safe_write_cell(ws, 'H3', data.get('jha_hazard1', ''))
        safe_write_cell(ws, 'I3', data.get('jha_control1', ''))
        
        # Row 5: JHA Step 2
        safe_write_cell(ws, 'G5', data.get('jha_step2', ''))
        safe_write_cell(ws, 'H5', data.get('jha_hazard2', ''))
        safe_write_cell(ws, 'I5', data.get('jha_control2', ''))
        
        # ============================================
        # SITE LOCATIONS - Multiple Sites
        # ============================================
        # These are in columns J-M (Site ID, Anchor ID, Safety Officer, Project In-Charge)
        sites = data.get('sites', [])
        site_row_start = 3  # Row 3 is where site data starts in the template
        
        for i, site in enumerate(sites):
            if i >= 10:  # Limit to 10 sites
                break
            row = site_row_start + i
            # Check if we need to clear previous data first
            safe_write_cell(ws, f'J{row}', site.get('site_id', ''))
            safe_write_cell(ws, f'K{row}', site.get('anchor_id', ''))
            safe_write_cell(ws, f'L{row}', site.get('safety_officer', data.get('safety_officer', 'Ronnie Alvin Chiu')))
            safe_write_cell(ws, f'M{row}', site.get('project_in_charge', data.get('person_in_charge', 'John Carlo Rabanes')))
            safe_write_cell(ws, f'N{row}', site.get('worker_name', ''))
        
        # ============================================
        # JOB HAZARD ASSESSMENT (Bottom table)
        # ============================================
        jha_steps = data.get('jha_steps', [])
        row_start = 48  # Starting row for JHA table
        
        for i, step in enumerate(jha_steps):
            if i >= 10:
                break
            row = row_start + i
            safe_write_cell(ws, f'A{row}', step.get('step', ''))
            safe_write_cell(ws, f'B{row}', step.get('hazard', ''))
            safe_write_cell(ws, f'D{row}', step.get('controls', ''))
        
        # ============================================
        # LIST OF WORKERS
        # ============================================
        workers = data.get('workers', [])
        worker_row_start = 44  # Starting row for workers list
        
        # Split workers into two columns (A and B) like in the template
        for i, worker in enumerate(workers):
            if i >= 16:  # Limit to 16 workers (8 per column)
                break
            row = worker_row_start + (i // 2)
            col = 'A' if i % 2 == 0 else 'B'
            safe_write_cell(ws, f'{col}{row}', worker)
        
        # ============================================
        # ACKNOWLEDGEMENT
        # ============================================
        safe_write_cell(ws, 'B51', data.get('prepared_by', 'RONNIE ALVIN-CHIU'))
        safe_write_cell(ws, 'C51', data.get('noted_by', 'John Carlo Rabanes'))
        
        # Approved status - check if YES or NO
        approved = data.get('approved_status', 'YES')
        if approved == 'YES':
            safe_write_cell(ws, 'E52', 'X')  # Mark YES
        else:
            safe_write_cell(ws, 'F52', 'X')  # Mark NO
            
        safe_write_cell(ws, 'G52', data.get('safety_officer_approval', 'PTG MIDC O&M Regional Manager'))
        
        # ============================================
        # Save the workbook
        # ============================================
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
        
        # Based on your template
        sub_contractor = st.text_input("Name of Sub Contractor", value="DNA")
        requesting_vendor = st.text_input("Requesting Vendor", value="NOKIA")
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
        
        # Site locations table
        num_sites = st.number_input("Number of sites", min_value=0, max_value=10, value=2)
        
        sites = []
        site_headers = ['Site ID', 'Anchor ID/PLA No.', 'Safety Officer', 'Project In-Charge', 'Name of Worker/Designation']
        
        # Create columns for site data
        for i in range(num_sites):
            with st.expander(f"Site {i+1}", expanded=i==0):
                col_s1, col_s2, col_s3 = st.columns(3)
                with col_s1:
                    site_id = st.text_input(f"Site ID {i+1}", value="MIN624" if i==0 else "MIN779")
                    safety_officer_site = st.text_input(f"Safety Officer {i+1}", value="Ronnie Alvin Chiu")
                with col_s2:
                    anchor_id = st.text_input(f"Anchor ID/PLA No. {i+1}", value="")
                    project_in_charge_site = st.text_input(f"Project In-Charge {i+1}", value="John Carlo Rabanes")
                with col_s3:
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
        'jha_step1': jha_step1,
        'jha_hazard1': jha_hazard1,
        'jha_control1': jha_control1,
        'jha_step2': jha_step2,
        'jha_hazard2': jha_hazard2,
        'jha_control2': jha_control2,
        'sites': sites,
        'jha_steps': jha_steps,
        'workers': workers,
        'prepared_by': prepared_by,
        'noted_by': noted_by,
        'approved_status': approved_by,
        'safety_officer_approval': safety_officer_approval
    }
    
    st.markdown("---")
    
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
                            'Sites': len(sites),
                            'JHA Entries': len(jha_steps),
                            'Workers': len(workers)
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
        1. Fill in project details
        2. Add site locations with worker designations
        3. Add JHA entries as needed
        4. List all workers
        5. Click 'Generate Excel File'
        6. Download the completed file
        """)
        
        st.header("📋 Template Requirements")
        st.markdown("""
        - Template file: `HSWP_template.xlsx`
        - Must be in the repository root
        - Original template is preserved
        - Only specified fields are modified
        """)
        
        st.header("📌 Modified Fields")
        st.markdown("""
        - Project Details (all fields)
        - JHA Assessment (top section)
        - Site Locations (with worker designations)
        - JHA Table (bottom section)
        - List of Workers (two columns)
        - Acknowledgement section
        """)

if __name__ == "__main__":
    main()
