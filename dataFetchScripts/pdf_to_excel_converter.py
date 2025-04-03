import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

def clean_text(text):
    """
    Clean extracted text by removing extra whitespace while preserving line breaks
    """
    # Split into lines to preserve formatting
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        # Remove multiple spaces within each line
        cleaned_line = re.sub(r'\s+', ' ', line)
        # Remove spaces before punctuation
        cleaned_line = re.sub(r'\s+([.,;)])', r'\1', cleaned_line)
        if cleaned_line.strip():  # Only add non-empty lines
            cleaned_lines.append(cleaned_line)
    
    # Join lines back together, preserving original line breaks
    return '\n'.join(cleaned_lines)

def clean_objectives_text(text):
    """
    Clean objectives text by removing standard headers while preserving additional information and formatting
    """
    standard_headers = [
        r'Mål för projektet[\s:]*',
        r'Mål för samverkansprojektet[\s:]*',
        r'Objective for the project[\s:]*',
        r'Collaboration project objectives[\s:]*'
    ]
    
    # Split into lines to preserve formatting
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        cleaned_line = line
        for header in standard_headers:
            cleaned_line = re.sub(f'^{header}', '', cleaned_line, flags=re.IGNORECASE)
        if cleaned_line.strip():  # Only add non-empty lines
            cleaned_lines.append(cleaned_line)
    
    # Join lines back together, preserving original line breaks
    return '\n'.join(cleaned_lines)

def extract_section(text, start_pattern, end_pattern):
    """
    Extract a section from text using start and end patterns, capturing everything between the headers
    """
    try:
        # Find all matches for both start and end patterns
        start_matches = list(re.finditer(start_pattern, text, re.IGNORECASE | re.MULTILINE))
        end_matches = list(re.finditer(end_pattern, text, re.IGNORECASE | re.MULTILINE))
        
        if not start_matches:
            return ""
        
        # Get the last start match (in case there are multiple)
        start_match = start_matches[-1]
        start_pos = start_match.end()
        
        # Find the first end match that comes after our start match
        content = ""
        for end_match in end_matches:
            if end_match.start() > start_pos:
                content = text[start_pos:end_match.start()].strip()
                break
        
        if not content and end_matches:
            # If we didn't find an end match after our start, something might be wrong
            st.write(f"Debug - Possible section overlap. Start pos: {start_pos}, End matches: {[m.start() for m in end_matches]}")
            
        return content
        
    except Exception as e:
        st.warning(f"Error in extract_section: {str(e)}")
        return ""

def extract_project_data(text):
    """
    Extract project data from text using specific headers
    """
    data = {
        "Insatsområde som projektet adresserar": "",
        "Projektets titel": "",
        "Mål för projektet": "",
        "Sammanfattning": "",
        "Koordinerande projektpart": "",
        "Övriga projektparter": "",
        "Totalt budgeterad kostnad för projektet": "",
        "Totalt sökt bidrag": ""
    }
    
    # First, find all section positions to ensure correct ordering
    section_positions = []
    for section in data.keys():
        if section in ["Totalt budgeterad kostnad för projektet", "Totalt sökt bidrag"]:
            continue
            
        pattern = {
            "Insatsområde som projektet adresserar": r'Insatsområde som projektet adresserar|Focus area of the project',
            "Projektets titel": r'Projektets titel|Project title',
            "Mål för projektet": r'Mål för projektet|Mål för samverkansprojektet|Objective for the project|Collaboration project objectives',
            "Sammanfattning": r'Sammanfattning|Summary',
            "Koordinerande projektpart": r'Koordinerande projektpart|Coordinator organization',
            "Övriga projektparter": r'Övriga projektparter|Other project parties'
        }[section]
        
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            section_positions.append((match.start(), section))
    
    # Sort sections by their position in the text
    section_positions.sort()
    
    # Extract content between sections
    for i, (pos, section) in enumerate(section_positions):
        start_pattern = {
            "Insatsområde som projektet adresserar": r'(?:Insatsområde som projektet adresserar|Focus area of the project)(?:\s*\(.*?\))?[\s:]*',
            "Projektets titel": r'(?:Projektets titel|Project title)(?:\s*\(.*?\))?[\s:]*',
            "Mål för projektet": r'(?:Mål för projektet|Mål för samverkansprojektet|Objective for the project|Collaboration project objectives)(?:\s*\(.*?\))?[\s:]*',
            "Sammanfattning": r'(?:Sammanfattning|Summary)(?:\s*\(.*?\))?[\s:]*',
            "Koordinerande projektpart": r'(?:Koordinerande projektpart|Coordinator organization)(?:\s*\(.*?\))?[\s:]*',
            "Övriga projektparter": r'(?:Övriga projektparter|Other project parties)(?:\s*\(.*?\))?[\s:]*'
        }[section]
        
        # End pattern is the start of the next section, or a general end pattern for the last section
        end_pattern = r'(?:Totalt budgeterad|Total budgeted)'
        if i < len(section_positions) - 1:
            next_section = section_positions[i + 1][1]
            end_pattern = {
                "Insatsområde som projektet adresserar": r'(?:Projektets titel|Project title)',
                "Projektets titel": r'(?:Mål för projektet|Mål för samverkansprojektet|Objective for the project|Collaboration project objectives)',
                "Mål för projektet": r'(?:Sammanfattning|Summary)',
                "Sammanfattning": r'(?:Koordinerande projektpart|Coordinator organization)',
                "Koordinerande projektpart": r'(?:Övriga projektparter|Other project parties)',
                "Övriga projektparter": r'(?:Totalt budgeterad|Total budgeted)'
            }[section]
        
        content = extract_section(text, start_pattern, end_pattern)
        
        # Debug output
        st.write(f"Debug - Section: {section}")
        st.write(f"Debug - Content length: {len(content)}")
        if content:
            st.write(f"Debug - First 50 chars: {content[:50]}")
        
        if content.strip():
            # Clean the content
            lines = content.split('\n')
            cleaned_lines = []
            for line in lines:
                cleaned_line = re.sub(r'\s+', ' ', line).strip()
                if cleaned_line:
                    cleaned_lines.append(cleaned_line)
            data[section] = '\n'.join(cleaned_lines)
    
    # Extract budget information using specific patterns
    budget_patterns = [
        (r'(?:Totalt budgeterad kostnad för projektet|Total budgeted cost)(?:\s*\(.*?\))?:\s*([0-9\s,\.]+(?:\s*(?:MSEK|SEK))?)',
         "Totalt budgeterad kostnad för projektet"),
        (r'(?:Totalt sökt bidrag|Total requested contribution)(?:\s*\(.*?\))?:\s*([0-9\s,\.]+(?:\s*(?:MSEK|SEK))?)',
         "Totalt sökt bidrag")
    ]
    
    for pattern, key in budget_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data[key] = match.group(1).strip()
    
    return data

def convert_pdf_to_excel(pdf_file):
    """
    Convert PDF file to Excel format with enhanced extraction
    """
    try:
        # Read PDF file
        with pdfplumber.open(pdf_file) as pdf:
            st.write(f"Total pages in PDF: {len(pdf.pages)}")
            
            # Store all project data
            all_projects = []
            
            # Process each page
            for page_num, page in enumerate(pdf.pages, 1):
                st.write(f"Processing page {page_num}...")
                
                # Extract text from page
                text = page.extract_text()
                
                # Skip pages without project information (check for both Swedish and English variants)
                if not any(header in text for header in ["Projektsammanfattning", "Project summary"]):
                    continue
                
                # Extract project data from text
                project_data = extract_project_data(text)
                
                # Debug information for empty sections
                empty_sections = [key for key, value in project_data.items() if not value.strip()]
                if empty_sections:
                    st.warning(f"Empty sections on page {page_num}: {', '.join(empty_sections)}")
                
                # Only add projects that have at least some data
                if any(value.strip() for value in project_data.values()):
                    # Debug information
                    st.write(f"Project title found on page {page_num}: {project_data['Projektets titel'][:100]}...")
                    st.write(f"Mål för projektet length: {len(project_data['Mål för projektet'])}")
                    st.write(f"Sammanfattning length: {len(project_data['Sammanfattning'])}")
                    
                    # Print a sample of the Sammanfattning content for debugging
                    if project_data['Sammanfattning']:
                        st.write(f"Sammanfattning sample: {project_data['Sammanfattning'][:100]}...")
                    
                    all_projects.append(project_data)
                else:
                    st.warning(f"No data could be extracted from page {page_num}")
                
            if not all_projects:
                st.error("No project data found in the PDF file.")
                return None
            
            # Create DataFrame from all projects
            df = pd.DataFrame(all_projects)
            st.write(f"Found {len(df)} projects")
            
            return df
            
    except Exception as e:
        st.error(f"Error processing PDF: {str(e)}")
        st.exception(e)
        return None

def main():
    st.title("PDF to Excel Converter")
    st.write("Upload a PDF file containing project information to convert it to Excel format.")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
    
    if uploaded_file is not None:
        # Display file details
        st.write("File details:")
        st.write(f"Filename: {uploaded_file.name}")
        st.write(f"File size: {uploaded_file.size / 1024:.2f} KB")
        
        # Convert button
        if st.button("Convert to Excel"):
            with st.spinner("Converting PDF to Excel..."):
                # Convert PDF to DataFrame
                df = convert_pdf_to_excel(uploaded_file)
                
                if df is not None:
                    # Create Excel file in memory
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Projects')
                        
                        # Auto-adjust column widths
                        worksheet = writer.sheets['Projects']
                        for idx, col in enumerate(df.columns):
                            max_length = max(
                                df[col].astype(str).apply(len).max(),
                                len(col)
                            ) + 2
                            worksheet.set_column(idx, idx, min(max_length, 100))
                    
                    # Get the Excel file from memory
                    excel_data = output.getvalue()
                    
                    # Create download button
                    st.download_button(
                        label="Download Excel file",
                        data=excel_data,
                        file_name=f"{uploaded_file.name.replace('.pdf', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Display preview of the data
                    st.write("Preview of the data:")
                    st.dataframe(df)
                    
                    # Display some statistics
                    st.write("Data Statistics:")
                    st.write(f"Total projects found: {len(df)}")
                    for col in df.columns:
                        non_empty = df[df[col].str.strip() != ''].shape[0]
                        st.write(f"{col}: {non_empty} non-empty values")

if __name__ == "__main__":
    main() 