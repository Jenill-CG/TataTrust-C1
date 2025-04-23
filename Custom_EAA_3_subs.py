import streamlit as st
import pandas as pd
import numpy as np
import zipfile
import os
import tempfile
import base64
import io
import xlsxwriter
from fpdf import FPDF
import streamlit_pdf_viewer as pdf_viewer
from streamlit_folium import st_folium
import folium
import plotly.express as px
import streamlit.components.v1 as components

# Define the parameter descriptions
parameter_descriptions = {
    'A1': "School + Grade + Student",
    'A2': "Block + School + Grade + Student",
    'A3': "District + School + Grade + Student",
    'A4': "Partner + School + Grade + Student",
    'A5': "District + Block + School + Grade + Student",
    'A6': "Partner + Block + School + Grade + Student",
    'A7': "Partner + District + School + Grade + Student",
    'A8': "Partner + District + Block + School + Grade + Student",
    'A9': "Partner + group + School + Grade + Student"
}

# Define the new mapping for parameter sets
parameter_mapping = {
    'A1': "School_ID,Grade,student_no",
    'A2': "Block_ID,School_ID,Grade,student_no",
    'A3': "District_ID,School_ID,Grade,student_no",
    'A4': "Partner_ID,School_ID,Grade,student_no",
    'A5': "District_ID,Block_ID,School_ID,Grade,student_no",
    'A6': "Partner_ID,Block_ID,School_ID,Grade,student_no",
    'A7': "Partner_ID,District_ID,School_ID,Grade,student_no",
    'A8': "Partner_ID,District_ID,Block_ID,School_ID,Grade,student_no",
    'A9': "Partner_ID,group,School_ID,Grade,student_no"
}

# Dropdown for selecting file naming format
naming_options = {
    "School Name + Block Name": "{school_name}_{block_name}",
    "School Name + District Name": "{school_name}_{district_name}",
    "School Name + Grade": "{school_name}_Grade{grade}"
}

def generate_custom_id(row, params):
    params_split = params.split(',')
    custom_id = []
    for param in params_split:
        if param in row and pd.notna(row[param]):
            value = row[param]
            if isinstance(value, float) and value % 1 == 0:
                value = int(value)
            custom_id.append(str(value))
    return ''.join(custom_id)

def process_data(uploaded_file, partner_id, buffer_percent, grade,group,district_digits, block_digits, school_digits, student_digits, selected_param):
    data = pd.read_excel(uploaded_file)
    # Check for duplicate School_IDs
    if data['School_ID'].duplicated().any():
        raise ValueError("Duplicate School_ID found in the uploaded file. Please ensure each School_ID is unique.")
    
    unique_school_count = data['School_ID'].nunique()
    digit_count = len(str(unique_school_count))
    if digit_count > school_digits:
        school_digits = digit_count
    
    # Assign the Partner_ID directly
    data['Partner_ID'] = str(partner_id).zfill(len(str(partner_id)))  # Padding Partner_ID
    data['Grade'] = grade
    data['group'] = group
    # Assign unique IDs for District, Block, and School, default to "00" for missing values
    # data['School_udise'] = data['School_ID'].astype(str).str.zfill(12)
    data['School_udise'] = data['School_ID']
    data['District_ID'] = data['District'].apply(lambda x: str(data['District'].unique().tolist().index(x) + 1).zfill(district_digits) if x != "NA" else "0".zfill(district_digits))
    data['Block_ID'] = data['Block'].apply(lambda x: str(data['Block'].unique().tolist().index(x) + 1).zfill(block_digits) if x != "NA" else "0".zfill(block_digits))
    data['School_ID'] = data['School_ID'].apply(lambda x: str(data['School_ID'].unique().tolist().index(x) + 1).zfill(school_digits) if x != "NA" else "0".zfill(school_digits))
    # Calculate Total Students With Buffer based on the provided buffer percentage
    data['Total_Students_With_Buffer'] = np.floor(data['Total_Students'] * (1 + buffer_percent / 100))
    
    # Generate student IDs based on the calculated Total Students With Buffer
    def generate_student_ids(row):
        if pd.notna(row['Total_Students_With_Buffer']) and row['Total_Students_With_Buffer'] > 0:
            student_ids = [
                f"{row['School_ID']}{str(int(row['Grade'])).zfill(2)}{str(i).zfill(student_digits)}"
                for i in range(1, int(row['Total_Students_With_Buffer']) + 1)
            ]
            return student_ids
        return []
    data['Student_IDs'] = data.apply(generate_student_ids, axis=1)
    # Expand the data frame to have one row per student ID
    data_expanded = data.explode('Student_IDs')
    # Extract student number from the ID
    data_expanded['student_no'] = data_expanded['Student_IDs'].str[-student_digits:]
    # Use the selected parameter set for generating Custom_ID
    data_expanded['Custom_ID'] = data_expanded.apply(lambda row: generate_custom_id(row, parameter_mapping[selected_param]), axis=1)
    # Generate the additional Excel sheets with mapped columns (without the Gender column)
    data_mapped = data_expanded[['Custom_ID', 'Grade', 'School', 'School_ID', 'District', 'Block','group']].copy()
    data_original_mapped = data_expanded[['Custom_ID', 'Grade', 'School', 'School_udise', 'District', 'Block','group']].copy()
    data_mapped.columns = ['Roll_Number', 'Grade', 'School Name', 'School Code', 'District Name', 'Block Name','group']
    data_original_mapped.columns = ['Roll_Number', 'Grade', 'School Name', 'School Code', 'District Name', 'Block Name','group']
    # Generate Teacher_Codes sheet
    teacher_codes = data[['School', 'School_ID']].copy()
    teacher_codes.columns = ['School Name', 'School Code']
    return data_expanded, data_mapped, teacher_codes, data_original_mapped

def download_link(df, filename, link_text):
    towrite = io.BytesIO()
    with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" class="download-link"><img src="https://img.icons8.com/material-outlined/24/000000/download.png" class="download-icon"/> {link_text}</a>'

# Function to create the attendance list PDF
def create_attendance_pdf(pdf, column_widths, column_names, image_path, info_values, df, format_option):
    pdf.add_page()

    # Set top margin to 2O mm
    pdf.set_top_margin(20)
    pdf.set_auto_page_break(auto=True, margin=20)

    # Page width and margins
    page_width = 210  # A4 page width in mm
    margin_left = 9
    margin_right = 9
    available_width = page_width - margin_left - margin_right

    # Calculate total column width
    total_column_width = sum(column_widths[col] for col in column_names)

    # Scale column widths if necessary
    if total_column_width > available_width:
        scaling_factor = available_width / total_column_width
        column_widths = {col: width * scaling_factor for col, width in column_widths.items()}

    # Move to 20 mm from the top
    pdf.set_y(20)

    # Set the Font for the Title and Subtitle
    pdf.set_font('Arial', 'B', 7)

    # Calculate the Width of the Merged Cell
    merged_cell_width = sum(column_widths[col] for col in column_names)  # Total width based on scaled column widths

    # Add the Title and Subtitle in the Center
    pdf.cell(merged_cell_width, 12, '', border='LTR', ln=1, align='C')  # Create an empty cell with borders

    # Set the cursor position back to the beginning of the merged cell
    pdf.set_xy(pdf.get_x(), pdf.get_y() - 10)

    # Centered Title
    pdf.cell(merged_cell_width, 4, 'ATTENDANCE LIST', border=0, align='C', ln=2)

    # Centered Subtitle
    pdf.set_font('Arial', '', 3)
    pdf.cell(merged_cell_width, 1, '(PLEASE FILL ALL THE DETAILS IN BLOCK LETTERS)', border=0, align='C', ln=1)

    # Bottom border of the merged cell
    pdf.cell(merged_cell_width, 3, '', border='LBR', ln=1)  # Bottom border of the merged cell

    # Add the image in the top-right corner of the bordered cell
    pdf.image(image_path, x=pdf.get_x() + 160, y=pdf.get_y() - 8, w=15, h=5.5)  # Adjust position and size as needed

    # Add the additional information cell below the "ATTENDANCE LIST" cell
    pdf.set_font('Arial', 'B', 5)
    info_cell_width = merged_cell_width  # Width same as the merged title cell
    info_cell_height = 15  # Adjust height as needed
    pdf.cell(info_cell_width, info_cell_height, '', border='LBR', ln=1)
    pdf.set_xy(pdf.get_x(), pdf.get_y() - info_cell_height)  # Move back to the top of the cell

    # Add labels and fill values from the dictionary
    info_labels = {
        'DISTRICT': '',
        'BLOCK': '',
        'SCHOOL NAME': '',
        'CLASS': '',
        'SCHOOL CODE': ''
    }

    # Prioritize exact matches, but still allow partial matching
    for label in info_labels.keys():
        matched = False  # Flag to check if an exact match is found
        for key, value in info_values.items():
            # First check for an exact match (ignoring case)
            if label.lower() == key.lower():
                info_labels[label] = value
                matched = True
                break
        if not matched:
            # If no exact match is found, use partial matching for first 5 characters
            for key, value in info_values.items():
                if label[:5].lower() == key[:5].lower():
                    info_labels[label] = value
                    break

    # Width for the school name and date of assessment cells
    school_name_width = info_cell_width * 0.65  # 65% of the total width for the school name
    date_width = info_cell_width * 0.35         # 35% of the total width for the date of assessment

    # Add the DISTRICT, BLOCK, and other labels
    pdf.cell(info_cell_width, 3, f"DISTRICT : {info_labels['DISTRICT']}", border='LR', ln=1)
    pdf.cell(info_cell_width, 3, f"BLOCK : {info_labels['BLOCK']}", border='LR', ln=1)

    # Add the SCHOOL NAME
    pdf.cell(school_name_width, 3, f"SCHOOL NAME : {info_labels['SCHOOL NAME']}", border='L', ln=0)  # Left border only

    # Set a different font for the DATE OF ASSESSMENT
    pdf.set_font('Arial', 'B', 4.5)  # Set to Arial, Italic, size 5

    # Add the DATE OF ASSESSMENT on the right side
    pdf.cell(date_width, 3, "DATE OF ASSESSMENT : ______________                       ", border='R', ln=1, align='R')  # Right border only

    # Reset the font back to the original for the remaining labels
    pdf.set_font('Arial', 'B', 5)
    
    minusone = info_labels['CLASS']
    #minusone = info_labels['CLASS']+1

    # Add the CLASS and SECTION labels
    pdf.cell(info_cell_width, 3, f"CLASS : {minusone}", border='LR', ln=1)
    pdf.cell(info_cell_width, 3, f"SCHOOL CODE : {info_labels['SCHOOL CODE']}", border='LR', ln=1)

    # Draw a border around the table header
    pdf.set_font('Arial', 'B', 5)
    table_cell_height = 9

    # Add the Title and Subtitle in the Center
    if format_option == 'Pen Paper Assessment':
        # Add the Title and Subtitle for pen paper format
        pdf.cell(6, 4, '', border='LTR', align='C')
        pdf.cell(16, 4, '', border='LTR', align='C')
        pdf.cell(73, 4, '', border='LTR', align='C')
        pdf.cell(12, 4, '', border='LTR', align='C')
        pdf.cell(20, 4, '', border='LTR', align='C')
        pdf.cell(20, 4, '', border='LTR', align='C')
        pdf.cell(20, 4, '', border='LTR', align='C')
        pdf.cell(12, 4, '', border='LTR', align='C')  # End of the row

        pdf.ln(4)
        # First row of headers
        pdf.cell(6, 0.5, 'S.NO', border='LR', align='C')
        pdf.cell(16, 0.5, 'STUDENT ID', border='LR', align='C')
        pdf.cell(73, 0.5, 'STUDENT NAME', border='LR', align='C')
        pdf.cell(12, 0.5, 'GENDER', border='LR', align='C')
        pdf.cell(20, 0.5, 'SUBJECT 1', border='LR', align='C')
        pdf.cell(20, 0.5, 'SUBJECT 2', border='LR', align='C')
        pdf.cell(20, 0.5, 'SUBJECT 3', border='LR', align='C')
        pdf.cell(12, 0.5, 'SESSION', border='LR', align='C')  # End of the row

        # Move to the next line
        pdf.ln(0.5)

        # Second row of headers (merged cells)
        pdf.set_font("Arial",  size=5)
        pdf.cell(6, 4.5, '', border='LBR', align='C')  # Empty cell under S.NO
        pdf.cell(16, 4.5, '', border='LBR', align='C')  # Empty cell under STUDENT ID
        pdf.cell(73, 4.5, '', border='LBR', align='C')  # Empty cell under STUDENT NAME
        pdf.cell(12, 4.5, '', border='LBR', align='C')  # Empty cell under GENDER
        pdf.cell(20, 4.5, 'Present/Absent', border='LBR', align='C')  # Empty cell under SUBJECT 1
        pdf.cell(20, 4.5, 'Present/Absent', border='LBR', align='C')  # Empty cell under SUBJECT 2
        pdf.cell(20, 4.5, 'Present/Absent', border='LBR', align='C')  # Empty cell under SUBJECT 3
        pdf.cell(12, 4.5, '', border='LBR', align='C')  # Empty cell under SESSION
        pdf.ln(4.5)

    elif format_option == 'Digital Assessment':
        # Add the Title and Subtitle for digital format
        pdf.cell(6, 4, '', border='LTR', align='C')
        pdf.cell(15, 4, '', border='LTR', align='C')
        pdf.cell(72, 4, '', border='LTR', align='C')
        pdf.cell(12, 4, '', border='LTR', align='C')
        # pdf.cell(18, 4, '', border='LTR', align='C')
        pdf.cell(34, 4, '', border='LTR', align='C')
        pdf.cell(16, 4, '', border='LTR', align='C')
        # pdf.cell(12, 4, '', border='LTR', align='C')
        pdf.cell(24, 4, '', border='LTR', align='C')  # End of the row

        pdf.ln(4)
        # First row of headers
        pdf.cell(6, 0.5, 'S.NO', border='LR', align='C')
        pdf.cell(15, 0.5, 'STUDENT ID', border='LR', align='C')
        pdf.cell(72, 0.5, 'STUDENT NAME', border='LR', align='C')
        pdf.cell(12, 0.5, 'GENDER', border='LR', align='C')
        # pdf.cell(18, 0.5, 'TAB ID', border='LR', align='C')
        pdf.cell(34, 0.5, 'HOME LANGUAGE', border='LR', align='C')
        pdf.cell(16, 0.5, 'MATH', border='LR', align='C')
        # pdf.cell(12, 0.5, 'SECTION', border='LR', align='C')
        pdf.cell(24, 0.5, 'LANGUAGE', border='LR', align='C')  # End of the row

        # Move to the next line
        pdf.ln(0.5)

        # Second row of headers (merged cells)
        pdf.set_font("Arial",  size=5)
        pdf.cell(6, 4.5, '', border='LBR', align='C')  # Empty cell under S.NO
        pdf.cell(15, 4.5, '', border='LBR', align='C')  # Empty cell under STUDENT ID
        pdf.cell(72, 4.5, '', border='LBR', align='C')  # Empty cell under STUDENT NAME
        pdf.cell(12, 4.5, '', border='LBR', align='C')  # Empty cell under GENDER
        # pdf.cell(18, 4.5, '', border='LBR', align='C')  # Empty cell under TAB ID
        pdf.cell(34, 4.5, '', border='LBR', align='C')  # Empty cell under SUBJECT 1
        pdf.cell(16, 4.5, '', border='LBR', align='C')  # Empty cell under SUBJECT 2
        # pdf.cell(12, 4.5, '', border='LBR', align='C')  # Empty cell under SECTION
        pdf.cell(24, 4.5, '', border='LBR', align='C')  # Empty cell under SESSION
        pdf.ln(4.5)
    # pdf.ln(4.5)

    # Table Rows (based on student_count)
    pdf.set_font('Arial', '', 6)
    student_count = info_values.get('student_count', 0)  # Use 0 if 'student_count' is missing or not found

    # Fill in the student IDs for the selected school code
    student_ids = df[df['School Code'] == info_values.get('School Code', '')]['STUDENT ID'].tolist()

    for i in range(student_count):
        # Fill in S.NO column
        pdf.cell(column_widths['S.NO'], table_cell_height, str(i + 1), border=1, align='C')

        # Fill in STUDENT ID column
        student_id = student_ids[i]
        pdf.cell(column_widths['STUDENT ID'], table_cell_height, str(student_id), border=1, align='C')

        # Fill in remaining columns with empty values
        for col_name in column_names[2:]:  # Skip first two columns
            pdf.cell(column_widths[col_name], table_cell_height, '', border=1, align='C')

        pdf.ln(table_cell_height)

def main():
    
    # Initialize session state
    if 'buttons_initialized' not in st.session_state:
        st.session_state['buttons_initialized'] = True
        st.session_state['generate_clicked'] = False
        st.session_state['download_data'] = None
        st.session_state['checkboxes_checked'] = False
        st.session_state['thank_you_displayed'] = False  # Initialize thank you state

    if st.session_state['thank_you_displayed']:
        st.markdown("""
            <div style='border: 1px solid #c3e6cb; padding: 15px; border-radius: 5px; background-color: #d4edda; color: #155724;'>
                <h2 style='text-align: center; color: #155724;'>😊 Thank You 😊</h2>
                <p style='text-align: center; font-size: 18px; color: #155724;'>We hope PDFs are meeting your expectations</p>
                <h3 style='text-align: center; color: #155724;'>We'd love to hear your feedback👇</h3>
                <p style='text-align: center;'><a href='https://forms.gle/jpeC9xmtzSBqSQhL9' target='_blank' style='color: #155724;'>Feedback form</a></p>
            </div>
        """, unsafe_allow_html=True)
        return



        # If the thank you message has already been displayed, show only the thank you message
    # if st.session_state['thank_you_displayed']:
    #     st.markdown("<h2 style='text-align: center; color: green;'>Thank You for using the Attendance Sheet Generator!</h2>", unsafe_allow_html=True)
    #     st.markdown("<p style='text-align: center; font-size: 18px;'>We hope the generated PDFs meet your expectations.</p>", unsafe_allow_html=True)
    #     st.markdown("<h3 style='text-align: center;'>We'd love to hear your feedback!</h3>", unsafe_allow_html=True)
    #     st.markdown("<p style='text-align: center;'>Please fill out our <a href='https://forms.gle/jpeC9xmtzSBqSQhL9' target='_blank'>feedback form</a>.</p>", unsafe_allow_html=True)
    #     return

    # Centered title
    css = """
    <style>
    .custom-header {
        font-size: 26px; /* Larger font size for prominence */
        color: #F0BF4C; /* Primary color for the text */
        text-align: right; /* Center the text */
        padding: 2px; /* Add padding around the text */
        background-color: #ffffff; /* Light background color */
        # border-radius: 10px; /* Rounded corners */
        # box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Subtle shadow */
        # text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.2); /* Text shadow for depth */
        margin-top: 2px; /* Add top margin */
        margin-bottom: 2px; /* Add bottom margin */
    }
    </style>
    """

    # Apply the custom CSS
    st.markdown(css, unsafe_allow_html=True)

    # Display the styled header
    # st.markdown("<div class='custom-header'>Welcome!</div>", unsafe_allow_html=True)
    # st.markdown("<div class='custom-header'>Tool for ID Generation</div>", unsafe_allow_html=True)
    st.markdown("<div class='custom-header' style='background-color: #F0BF4C; text-align: center;padding: 15px; font-size: 40px;font-weight: bold;color: #FFFFFF; border-radius: 10px; box-shadow: 0px 4px 8px rgba(0, 0, 0, 0.2);'>CGI's Custom ID Builder</div>", unsafe_allow_html=True)
    st.markdown("<div class='custom-header' style='font-size: 26px;color: #F0BF4C;text-align: right;padding: 2px; background-color: #ffffff; margin-top: 2px;margin-bottom: 2px;box-shadow: 0px 4px 8px rgba(0, 0, 0, 0);text-shadow: 1px 1px 2px rgba(0, 0, 0, 0);'>Generate unique IDs quickly and easily!</div>", unsafe_allow_html=True)

    # Data for the example table
    data = {
        'School_ID': [1001],
        'District': ['District A'],
        'Block': ['Block A'],
        'School': ['School A'],
        'Total_Students': [300]
    }
    df = pd.DataFrame(data)
    
    # Convert DataFrame to HTML
    html_table = df.to_html(index=False, border=0, classes='custom-table')
    
    # Custom CSS to style the table and the warning box
    css = """
    <style>
    .custom-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
        margin-top: 2px;
    }
    .custom-table th, .custom-table td {
        padding: 10px;
        text-align: center;
        border: 1px solid #ddd;
    }
    .custom-table th {
        background-color: #F4F4F4;
    }
    .download-link {
        color: green;
        text-decoration: none;
        font-weight: bold;
    }
    .download-link:hover {
        text-decoration: underline;
    }
    .download-icon {
        margin-right: 8px;
    }
    .warning-box {
        background-color: #FFFFE0;
        border: 1px solid #FFD700;
        padding: 10px;
        margin-top: 10px;
        border-radius: 5px;
    }
    </style>
    """
    
    # Display the text and table
    st.markdown(css, unsafe_allow_html=True)
    # st.warning("""Please rename your column headers as per input file structure shown""")
    st.markdown("➡️ Please modify input column headers to match the specified structure🗒️")
    st.markdown(html_table, unsafe_allow_html=True)
    
    st.info(
        """
        **Note:**
        
        - Please make sure that input values in each raw of "School_ID" column are UNIQUE.
        - This program will only accept a single sheet in the input and will not permit hidden sheets.
        """
    )
    
    #File uploader section
    st.markdown("➡️ Upload an Excel file")
    uploaded_file = st.file_uploader("Please upload an XLSX file that is less than 200MB in size",type=["xlsx"])
    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file)
        unique_school_count = data['School_ID'].nunique()
        school_digit_count = len(str(unique_school_count))
        unique_district_count = data['District'].nunique()
        district_digit_count = len(str(unique_district_count))
        unique_block_count = data['Block'].nunique()
        block_digit_count = len(str(unique_block_count))
        student_digit_count = len(str(max(data['Total_Students'])))


        
        # Centered and colored message
        st.markdown("<p style='text-align: center; color: green;font-size: 26px;'>File uploaded successfully 🥳</p>", unsafe_allow_html=True)
        
        col1, col2= st.columns([1,1])
        with col1:    
            run_default = st.checkbox("IDs with Default Settings")
        with col2:
            customize_id = st.checkbox("IDs with Customized Settings")

        # Checkboxes to select mode
        # run_default = st.checkbox("IDs with Default Settings")
        # customize_id = st.checkbox("IDs with Customized Settings")
        
        # Ensure only one checkbox is selected
        if run_default and customize_id:
            st.warning("Please select only one option.")
            return

        # Set checkboxes_checked to True if either checkbox is selected
        st.session_state['checkboxes_checked'] = run_default or customize_id
        
        if run_default:
            # Default parameters
            partner_id = 11
            col1, col2= st.columns([1,2])
            with col1:
                grade = st.number_input("➡️ Please provide required Grade", min_value=1, value=1)
            with col2:
                vb = st.write("")
            # grade = st.number_input("Grade", min_value=1, value=1)
            # buffer_percent = 0
            buffer_percent = 10.0
            district_digits = district_digit_count
            block_digits = block_digit_count
            school_digits = school_digit_count
            student_digits = student_digit_count
            selected_param = 'A4'  # Default parameter
        elif customize_id:
            # # Custom parameters
            # st.markdown("<p style='color: blue;'>Please provide required Values</p>", unsafe_allow_html=True)
            st.markdown("➡️ Please provide required Values", unsafe_allow_html=True)
            col1, col2, col3, col4 = st.columns([1,1,1,1])
            with col1:
                partner_id = st.number_input("Partner ID", min_value=12, value=12)
            with col2:
                buffer_percent = st.number_input("Buffer Percentage", min_value=0.0, value=10.0, format="%.2f")
                # buffer_percent =st.slider("Buffer Percentage",min_value=0.0,max_value=50.0,value=(0.0, 50.0),step=5.0)
            with col3:        
                grade = st.number_input("Grade", min_value=1, value=1)
            with col4: 
                group = st.number_input("Group Id", min_value=1, value=1)

            # partner_id = st.number_input("Partner ID", min_value=12, value=12)

            # col1, col2 = st.columns([1, 3])
            # with col1:
            # # Select slider with reduced width placed in the first narrow column
            #     st.write("Enter values")
            # #value = st.select_slider("Select a value",options=[0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50],value=50)
            # with col2:
            #     buffer_percent =st.slider("Buffer Percentage",min_value=0,max_value=50,value=(0, 50),step=5)
            

            #buffer_percent = st.number_input("Buffer Percentage", min_value=0.0, value=0.0, format="%.2f")
            
            #buffer_percent =st.slider("Buffer Percentage",min_value=0,max_value=50,value=(0, 50),step=5)
            #buffer_percent =st.radio("Buffer Percentage",options=[0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50])
            #buffer_percent = st.select_slider("Buffer Percentage",options=[0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50],value=0)

            # grade = st.number_input("Grade", min_value=1, value=1)
                
            # Message in blue color above District ID Digits
            # st.markdown("""➡️ Please provide required Digits <span style='color: blue;'>(Please select more than "minimum required value")</span></p>""", unsafe_allow_html=True)
            st.markdown("➡️ Please provide required Digits", unsafe_allow_html=True)

            col1, col2,col3 , col4 = st.columns([1,1,1,1])
            with col1:
                district_digits = st.number_input("District ID Digits", min_value=district_digit_count, value=2)
            with col2:
                block_digits = st.number_input("Block ID Digits", min_value=block_digit_count, value=2)
            with col3:
                school_digits = st.number_input("School ID Digits", min_value=school_digit_count, value=5)
            with col4:
                student_digits = st.number_input("Student ID Digits", min_value=student_digit_count, value=5)


            # district_digits = st.number_input("District ID Digits", min_value=district_digit_count, value=2)
            # block_digits = st.number_input("Block ID Digits", min_value=block_digit_count, value=2)
            # school_digits = st.number_input("School ID Digits", min_value=school_digit_count, value=5)
            # student_digits = st.number_input("Student ID Digits", min_value=student_digit_count, value=5)
            
            # Display parameter descriptions directly in selectbox
            parameter_options = list(parameter_descriptions.values())
            # st.markdown("""<style>.custom-selectbox-label {color: blue; margin: 0;}</style><p class='custom-selectbox-label'>Please Select Parameter Set for Desired Combination of Student IDs</p>""",unsafe_allow_html=True)
            st.markdown("""➡️ Please select the Combination""",unsafe_allow_html=True)
            selected_description = st.selectbox("Desired combination for Student IDs", parameter_options)
    
            # Get the corresponding parameter key
            selected_param = list(parameter_descriptions.keys())[parameter_options.index(selected_description)]

            if school_digit_count > school_digits:
                school_digits = school_digit_count
            if district_digit_count > school_digits:
                district_digits = district_digit_count
            if block_digit_count > school_digits:
                block_digits = block_digit_count

            # Create the format string based on selected_param
            param_description = parameter_descriptions[selected_param]
            format_parts = param_description.split(' + ')

            format_string = ' '.join([f"{'X' * (school_digits if 'School' in part else 
            block_digits if 'Block' in part else 
            district_digits if 'District' in part else 
            len(str(grade)) if 'Grade' in part else 
            # len(str(group)) if 'group' in part else 
            len(str(partner_id)) if 'Partner' in part else 
            student_digits)}" for part in format_parts])

            school_format = 'X' * school_digits

            # Display the ID format with a smaller font size
            st.markdown(f"<p style='color: blue; font-size: small;'>Your ID format would be: {format_string}</p>", unsafe_allow_html=True)
            # Display the School Code format based on the selected parameter


            # replace with above var
            st.markdown(f"<p style='color: blue; font-size: small;'>Your School Code format would be: {school_format}</p>", unsafe_allow_html=True)
        
        # Generate button action
        if st.session_state['checkboxes_checked']:
            if st.button("Generate IDs"):
                if uploaded_file is not None:
                    try:
                        # Process the uploaded file
                        expanded_data, mapped_data, teacher_codes, data_original = process_data(
                            uploaded_file,
                            partner_id,
                            buffer_percent,
                            grade,
                            group,
                            district_digits,
                            block_digits,
                            school_digits,
                            student_digits,
                            selected_param
                        )
                        # Update session state with generated data
                        st.session_state['download_data'] = (expanded_data, mapped_data, teacher_codes, data_original)
                        st.session_state['generate_clicked'] = True
                    except Exception as e:
                        st.error(f"Error processing file: {e}")

    # Download buttons after IDs are generated
    if st.session_state['generate_clicked'] and st.session_state['download_data'] is not None:
        expanded_data, mapped_data, teacher_codes, data_original = st.session_state['download_data']

        df1 = data_original
        # Define possible variations of 'Student ID' column names
        student_id_variations = ['STUDENT ID', 'STUDENT_ID', 'ROLL_NUMBER', 'Roll_Number', 'Roll Number']
        # Identify the actual column name from the variations
        student_id_column = None
        for variation in student_id_variations:
            if variation in df1.columns:
                student_id_column = variation
                break

        if student_id_column is None:
            raise ValueError("No recognized student ID column found in the data")

        class_variations = ['CLASS', 'Class', 'GRADE', 'Grade']
        # Identify the actual column name from the variations
        class_column = None
        for variation in class_variations:
            if variation in df1.columns:
                class_column = variation
                break
        if class_column is None:
            raise ValueError("No recognized student ID column found in the data")

        # Standardize column name to 'STUDENT_ID'
        df = df1.rename(columns={student_id_column: 'STUDENT ID', class_column: 'CLASS'})
        # Process data
        grouping_columns = [col for col in df.columns if col not in ['STUDENT ID', 'Gender'] and df[col].notna().any()]
        grouped = df.groupby(grouping_columns).agg(student_count=('STUDENT ID', 'nunique')).reset_index()

        if 'CLASS' in grouped.columns and grouped['CLASS'].astype(str).str.contains('\D').any():
            grouped['CLASS'] = grouped['CLASS'].astype(str).str.extract('(\d+)')

        result = grouped.to_dict(orient='records')

        # KPI Cards
        css = """
        <style>
        .custom-subheader {
            font-size: 24px; /* Larger font size */
            font-weight: bold; /* Bold font weight */
            color: #007bff; /* Primary color for the text */
            text-align: center; /* Center the text */
            padding: 10px; /* Add padding around the text */
            border-bottom: 2px solid #007bff; /* Add a border bottom */
            margin-top: 20px; /* Add top margin */
            margin-bottom: 20px; /* Add bottom margin */
            background-color: #f1f1f1; /* Light background color */
            border-radius: 5px; /* Rounded corners */
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); /* Subtle shadow */
        }
        </style>
        """

        # Apply the custom CSS
        st.markdown(css, unsafe_allow_html=True)

        # Display the styled subheader
        st.markdown("<div class='custom-subheader'>Your Summary</div>", unsafe_allow_html=True)

        # Calculating KPIs
        num_students = len(df['STUDENT ID'].unique())
        num_schools = df['School Code'].nunique() if 'School Code' in df.columns else 0
        num_blocks = df['Block Name'].nunique() if 'Block Name' in df.columns else 0
        num_districts = df['District Name'].nunique() if 'District Name' in df.columns else 0
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Number of Students", num_students)
        with col2:
            st.metric("Number of Schools", num_schools)
        with col3:
            st.metric("Number of Blocks", num_blocks)
        with col4:
            st.metric("Number of Districts", num_districts)

        
        # Download button for full data with Custom_IDs and Student_IDs
        #st.markdown(download_link(expanded_data, "full_data.xlsx", "Download Full Data (with Custom_IDs and Student_IDs)"), unsafe_allow_html=True)
        
        # Download button for mapped data
        st.markdown(download_link(mapped_data, "Student_Ids.xlsx", "Download Student IDs"), unsafe_allow_html=True)

        # Download button for teacher codes
        st.markdown(download_link(teacher_codes, "School_Codes.xlsx", "Download School Codes"), unsafe_allow_html=True)

    # if st.session_state['mapped_data'] is not None:
        # Centered title
        css = """
        <style>
        .custom-header {
            font-size: 36px; /* Larger font size for prominence */
            font-weight: bold; /* Bold font weight */
            # color: #20c997; /* Primary color for the text */
            color: #F0BF4C;
            text-align: center; /* Center the text */
            padding: 20px; /* Add padding around the text */
            background-color: #f8f9fa; /* Light background color */
            border-radius: 10px; /* Rounded corners */
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Subtle shadow */
            text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.2); /* Text shadow for depth */
            margin-top: 30px; /* Add top margin */
            margin-bottom: 30px; /* Add bottom margin */
        }
        </style>
        """

        # Apply the custom CSS
        st.markdown(css, unsafe_allow_html=True)

        # Display the styled header
        st.markdown("<div class='custom-header'>Attendance Sheet Generator</div>", unsafe_allow_html=True)

        image_path = "https://raw.githubusercontent.com/Jenill-CG/first/main/cg.png"

        # Choose between pen paper format or digital
        format_option = st.radio("➡️ Choose the format for the attendance sheet", ('Digital Assessment','Pen Paper Assessment'))
        
        # Number of columns and column names for the table based on the selected format
        if format_option == 'Pen Paper Assessment':
            column_names = ['S.NO', 'STUDENT ID', 'STUDENT NAME', 'GENDER', 'SUBJECT 1', 'SUBJECT 2', 'SUBJECT 3', 'SESSION']
            column_widths = {
                'S.NO': 6,
                'STUDENT ID': 16,
                'STUDENT NAME': 73,
                'GENDER': 12,
                'SUBJECT 1': 20,
                'SUBJECT 2': 20,
                'SUBJECT 3': 20,
                'SESSION': 12
            }
        else:
            column_names = ['S.NO', 'STUDENT ID', 'STUDENT NAME', 'GENDER', 'HOME LANGUAGE', 'MATH',  'LANGUAGE']
            column_widths = {
                'S.NO': 6,
                'STUDENT ID': 15,
                'STUDENT NAME': 72,
                'GENDER': 12,
                # 'TAB ID': 18,
                'HOME LANGUAGE': 34,
                'MATH': 16,
                # 'SECTION': 12,
                'LANGUAGE': 24
            }

        # selected_option = st.selectbox("➡️ Choose your file naming format", list(naming_options.keys()))
        # filename_template = naming_options[selected_option]
        
        col1, col2= st.columns([2,1])
        with col1:
            selected_option = st.selectbox("➡️ Choose your file naming format", list(naming_options.keys()))
        with col2:
            vb2 = st.write("")

        filename_template = naming_options[selected_option]

        if st.button("Click here to Generate PDFs"):
            with tempfile.TemporaryDirectory() as tmp_dir:
                pdf_paths = []
                preview_pdf_path = None  # To store the path of the first PDF
        
                # Create folders for districts
                district_folders = {}
                for record in result:
                    district_name = record.get('District Name', 'default_district')
                    if district_name not in district_folders:
                        district_folder = os.path.join(tmp_dir, district_name)
                        os.makedirs(district_folder, exist_ok=True)
                        district_folders[district_name] = district_folder
        
                for index, record in enumerate(result):
                    school_name = record.get('School Name', 'default_school').replace('/', '|')
                    district_name = record.get('District Name', 'default_district').replace('/', '|')
                    block_name = record.get('Block Name', 'default_block').replace('/', '|')
                    grade = record.get('CLASS', 'default_grade')
        
                    file_name = filename_template.format(school_name=school_name, district_name=district_name, block_name=block_name, grade=grade)
        
                    pdf = FPDF(orientation='P', unit='mm', format='A4')
                    pdf.set_left_margin(15)
                    pdf.set_right_margin(15)
        
                    create_attendance_pdf(pdf, column_widths, column_names, image_path, record, df, format_option)
        
                    # Save the PDF in the appropriate district folder
                    pdf_path = os.path.join(district_folders[district_name], f'{file_name}.pdf')
                    pdf.output(pdf_path)
                    pdf_paths.append(pdf_path)
        
                    if index == 0:  # Save the first PDF for preview
                        preview_pdf_path = pdf_path

                # Custom smaller header for PDF Preview
                st.markdown("""<h3 style='text-align: left; font-size:24px; color:#4CAF50;'>PDF Preview</h3>""",unsafe_allow_html=True)

                if preview_pdf_path:
                    # Read the PDF file as binary
                    with open(preview_pdf_path, "rb") as pdf_file:
                        pdf_data = pdf_file.read()
                        base64_pdf = base64.b64encode(pdf_data).decode('utf-8')
                        # Create a download link for the PDF
                        pdf_link = f'<a href="data:application/pdf;base64,{base64_pdf}" download="{os.path.basename(preview_pdf_path)}">Click here to download and view PDF</a>'
                        
                        # Display the link in Streamlit
                        st.markdown(pdf_link, unsafe_allow_html=True)
        
                # Create a zip file containing all district folders
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for district_name, folder_path in district_folders.items():
                        for foldername, _, filenames in os.walk(folder_path):
                            for filename in filenames:
                                filepath = os.path.join(foldername, filename)
                                # Preserve directory structure in ZIP file
                                arcname = os.path.relpath(filepath, tmp_dir)
                                zip_file.write(filepath, arcname)
        
                zip_buffer.seek(0)  # Reset buffer position
        
                # Provide download link for the zip file
                st.download_button(
                    label="Click to Download Zip File",
                    data=zip_buffer.getvalue(),
                    file_name="Attendance_sheets.zip",
                    mime="application/zip"
                )
                st.session_state['thank_you_displayed'] = True  # Set the thank you message state

if __name__ == "__main__":
    main()
