import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import os

st.set_page_config(page_title="KEA Report Generator", page_icon="📝", layout="wide")

st.title("📝 KEA COMPREHENSIVE SCHOOL")
st.subheader("Automated Student Assessment Report Generator")

# Check if the template exists in the GitHub repository folder
TEMPLATE_PATH = "template.docx"

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"⚠️ System Error: '{TEMPLATE_PATH}' not found in the repository. Please upload it to your GitHub files.")
else:
    st.success("✅ KEA Report Template loaded and ready.")

    # Now teachers only need to upload the Excel file!
    uploaded_excel = st.file_uploader("Upload Student Data Excel Sheet (.xlsx)", type=["xlsx"])

    if uploaded_excel:
        try:
            df = pd.read_excel(uploaded_excel)
            st.success("Excel data loaded successfully!")
            st.write("### Preview of Uploaded Student Data", df.head())
            
            if st.button("Generate Assessment Reports"):
                with st.spinner("Processing reports... Please wait."):
                    
                    zip_buffer = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for index, row in df.iterrows():
                            # Load the permanent template from the file path
                            doc = DocxTemplate(TEMPLATE_PATH)
                            
                            context = row.to_dict()
                            cleaned_context = {k: (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)) for k, v in context.items()}
                            
                            doc.render(cleaned_context)
                            
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            doc_io.seek(0)
                            
                            filename = f"{cleaned_context.get('Admin_No', index)}_{cleaned_context.get('Name', 'Student')}_Assessment_Report.docx"
                            filename = filename.replace("/", "-")
                            
                            zip_file.writestr(filename, doc_io.getvalue())
                    
                    zip_buffer.seek(0)
                    
                    st.success("🎉 All assessment reports generated successfully!")
                    st.download_button(
                        label="📥 Download All Reports (ZIP File)",
                        data=zip_buffer,
                        file_name="KEA_Comprehensive_School_Reports.zip",
                        mime="application/zip"
                    )
                    
        except Exception as e:
            st.error(f"An error occurred: {e}")
