import streamlit as st
import os
import tempfile
from openai import OpenAI
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from langchain_community.chat_models import ChatOpenAI
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
from langchain.text_splitter import RecursiveCharacterTextSplitter
from tqdm import tqdm
from tenacity import retry, stop_after_attempt, wait_exponential
import concurrent.futures
import tiktoken
import io
import zipfile

# Streamlit page config
st.set_page_config(layout='wide', page_title="Lease Synopsis Generator", page_icon="📄")

# OpenAI API setup
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def create_chat_llm():
    return ChatOpenAI(temperature=0.1, model="gpt-4")

chat_llm = create_chat_llm()

# Your prompt_template, chain, num_tokens_from_string, chunk_document, process_chunk, 
# generate_response, post_process_response, create_formatted_docx, consolidate_synopses, 
# and summarize_consolidated_synopsis functions remain the same

def process_uploaded_files(uploaded_files):
    reports_folder = tempfile.mkdtemp()
    all_synopses = []
    
    progress_bar = st.progress(0)
    for i, uploaded_file in enumerate(uploaded_files):
        document_text = extract_text_from_uploaded_file(uploaded_file)
        
        if document_text.strip():
            response = generate_response(document_text)
            output_file = os.path.join(reports_folder, f"synopsis_{uploaded_file.name}.docx")
            create_formatted_docx(response, output_file)
            all_synopses.append(response)
        
        progress_bar.progress((i + 1) / len(uploaded_files))
    
    # Generate consolidated report
    consolidated_report = consolidate_synopses(all_synopses)
    summarized_report = summarize_consolidated_synopsis(consolidated_report)
    create_formatted_docx(summarized_report, os.path.join(reports_folder, "consolidated_synopsis.docx"))
    
    return reports_folder

def extract_text_from_uploaded_file(uploaded_file):
    bytes_data = uploaded_file.read()
    doc = docx.Document(io.BytesIO(bytes_data))
    return '\n'.join([para.text for para in doc.paragraphs])

def create_download_zip(folder_path):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                zip_file.write(file_path, os.path.basename(file_path))
    return zip_buffer.getvalue()

# Streamlit UI
st.title("📄 Lease Synopsis Generator")
st.markdown("---")

uploaded_files = st.file_uploader("Upload .docx files", type="docx", accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully")
    if st.button("Generate Lease Synopsis"):
        with st.spinner("🔎 Generating lease synopsis..."):
            reports_folder = process_uploaded_files(uploaded_files)
            zip_file = create_download_zip(reports_folder)
            
            st.success("✅ Lease synopses generated successfully")
            
            st.download_button(
                label="Download Lease Synopses",
                data=zip_file,
                file_name="lease_synopses.zip",
                mime="application/zip"
            )
else:
    st.warning("⚠️ Please upload .docx files to generate lease synopses")

st.markdown("---")
st.markdown("Created with ❤️ by Weaver")
