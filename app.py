import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import os

st.set_page_config(page_title="KEA Report Generator", page_icon="📝", layout="wide")

st.title("📝 KEA COMPREHENSIVE SCHOOL")
st.subtitle("Automated Student Assessment Report Generator")
st.write("Upload your Report Card Word Template and Student Data Excel sheet to generate individual reports.")

---

# Sidebar Instructions
st.sidebar.header("Instructions")
st.sidebar.markdown("""
1. **Word Template:** Ensure your `.docx` file uses Jinja2 tags like `{{Name}}`, `{{Admin_No}}`, `{{Maths}}`, etc.
2. **Excel Data:** Column names in Excel must **exactly match** the tags inside your Word document.
3. **Click Generate:** Download all reports packaged in a single ZIP file.
""")

# File Uploaders
uploaded_template = st.file_uploader("1. Upload Report Template (.docx)", type=["docx"])
uploaded_excel = st.file_uploader("2. Upload Student Data (.xlsx)", type=["xlsx"])

if uploaded_template and uploaded_excel:
    try:
        # Read Excel Data
        df = pd.read_excel(uploaded_excel)
        st.success("Excel data loaded successfully!")
        
        # Preview Data
        st.write("### Preview of Uploaded Student Data", df.head())
        
        # Button to generate
        if st.button("Generate Assessment Reports"):
            with st.spinner("Processing reports... Please wait."):
                
                # Create an in-memory zip file
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    # Loop through each student record
                    for index, row in df.iterrows():
                        # Reload the template for each student
                        # Read template into bytes to allow repeating reads in memory
                        template_bytes = io.BytesIO(uploaded_template.getvalue())
                        doc = DocxTemplate(template_bytes)
                        
                        # Convert row to dictionary for context matching tags
                        context = row.to_dict()
                        
                        # Clean up data (e.g., handle NaN or floats for ranking/admin numbers)
                        cleaned_context = {k: (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)) for k, v in context.items()}
                        
                        # Render template with student data
                        doc.render(cleaned_context)
                        
                        # Save rendered document to an in-memory file
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        # Create a unique filename for each student
                        filename = f"{cleaned_context.get('Admin_No', index)}_{cleaned_context.get('Name', 'Student')}_Assessment_Report.docx"
                        filename = filename.replace("/", "-") # Prevent directory issues
                        
                        # Write to Zip
                        zip_file.writestr(filename, doc_io.getvalue())
                
                zip_buffer.seek(0)
                
                st.success("🎉 All assessment reports generated successfully!")
                
                # Download Button for Zip
                st.download_button(
                    label="📥 Download All Reports (ZIP File)",
                    data=zip_buffer,
                    file_name="KEA_Comprehensive_School_Reports.zip",
                    mime="application/zip"
                )
                
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload both the Word template and the Excel data sheet to proceed.")
