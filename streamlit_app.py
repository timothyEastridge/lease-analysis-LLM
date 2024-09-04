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

st.set_page_config(layout='wide', page_title="Lease Synopsis Generator", page_icon="📄")

# Custom CSS
st.markdown("""
<style>
    .stButton > button {
        width: 100%;
        background-color: #4CAF50;
        color: white !important;
        font-size: 18px;
        font-weight: bold;
        padding: 10px 24px;
        border-radius: 5px;
        border: none;
        cursor: pointer;
        transition: background-color 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #45a049;
    }
    .stTextInput > div > div > input {
        font-size: 16px;
    }
    h1 {
        color: #2C3E50;
        text-align: center;
        padding-bottom: 20px;
    }
    h2 {
        color: #34495E;
    }
    .fullWidth {
        width: 100%;
    }
    .reportview-container .main .block-container {
        max-width: 95%;
        padding-top: 5rem;
        padding-right: 1rem;
        padding-left: 1rem;
        padding-bottom: 5rem;
    }
</style>
""", unsafe_allow_html=True)

# OpenAI API key setup
if "openai" not in st.secrets or "api_key" not in st.secrets["openai"]:
    st.error("OpenAI API key not found in Streamlit secrets. Please add it to your app's secrets under [openai] api_key.")
    st.stop()

openai_api_key = st.secrets["openai"]["api_key"]
client = OpenAI(api_key=openai_api_key)

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def create_chat_llm():
    return ChatOpenAI(temperature=0.1, model="gpt-4", openai_api_key=openai_api_key)

chat_llm = create_chat_llm()

prompt_template = """
You are an expert in lease document analysis. Given a lease document or a portion of a lease document, extract the following information. Be concise and only return the information requested of you. If information is not found or not applicable, write "" instead of "N/A". Do not include duplicate or redundant information.

Location (e.g. 'Memorial City Towers, Ltd.'): 
Address:
Tenant Reference Name (doing business as):
Tenant Entity:
Guarantor: (keep this concise. For example, instead of Cenovus Energy Inc., a Canadian corporation just say Cenovus Energy Inc.)
Tenant's Notice Address (prior to occupancy):
Tenant's Notice Address (after occupancy):
Landlord's Notice Address (if mailed):
Landlord's Notice Address (if delivered):
Landlord's Payment Address:
Leased Premises:
Square Feet:
Commencement Date:
Expiration Date:
Extension Options:
Base Rent:
Operating Expenses: (add as much information as possible. for example: Tenant pays proportionate share of expenses (net lease), grossed up to reflect 100% occupancy. Management fee is 4% of rents. Commencing 04/01/25, expenses are capped at 6% accumulating and compounding amounts over FYE 03/31/25 amounts, except as attributable to insurance premiums/deductibles, increases in security due to staffing levels, janitorial or other costs which increase due to unionization, utilities and real estate tax/protest costs.)
Parking: (add as much info as possible. For example, Up to 24 non-reserved @ $50 per month, of which up to 5 may be reserved @ $100 per month. At any point during the term, T may convert an additional 2 non-reserved parking spaces into reserved parking spaces**Non-reserved parking charges abate from 10/01/24 - 02/16/30.)
Construction/Allowance:
Landlord's Relocation Rights: (add as much info as possible. for example, Landlord may relocate Tenant once during the term (except during first 2 years and last year of initial term and except during the  first or last year of extension option) upon 120 days prior written to another space in the building on 14th floor or higher, of a size between 100% to 110% of the premises at LL cost. Substitute premises to be improved with reasonably comparable or better quality leasehold improvements as existed in the premises. LL will provide T with at least 30 days access to the substitute premises after LL's tender  in order for T to install wiring, cabling, furniture, fixtures and equipment in the substitute premises at no cost to Tenant (Sec.3.3))
Tenant's Preferential Rights:
Termination Options:
Sign Rights:
Exclusive:
Use Restrictions on Landlord:
Build Restrictions on Landlord:
Off-site restrictions on Landlord:
Security Deposit:
Default Cure Period:
Holdover:
Broker/Commission:
Notice Address: (be sure to look for corporate address. for example, Leased PremisesWith a copy of all notices of default to the Guarantor at: Cenovus Energy Inc. 225 6 Avenue SW Calgary, AB T 2P 1N2 Attn: Director, Enterprise Compliance & Credit email: creditgroup@cenovus.com copy to: downstream.legal@cenovus.com)
Other Provisions:
Hazardous Material:
Insurance:
Tenant's Broker:
Special Provisions:

Document Text: {document_text}
"""

prompt = PromptTemplate(template=prompt_template, input_variables=["document_text"])
chain = LLMChain(llm=chat_llm, prompt=prompt)

def num_tokens_from_string(string: str, encoding_name: str = "cl100k_base") -> int:
    encoding = tiktoken.get_encoding(encoding_name)
    num_tokens = len(encoding.encode(string))
    return num_tokens

def chunk_document(text, max_tokens=4000, chunk_overlap=200):
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=max_tokens,
        chunk_overlap=chunk_overlap,
        length_function=lambda x: num_tokens_from_string(x)
    )
    return text_splitter.split_text(text)

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def process_chunk(chunk):
    try:
        response_dict = chain.invoke({"document_text": chunk})
        return response_dict.get('text', '').strip()
    except Exception as e:
        return f'Error in processing chunk: {str(e)}'

def generate_response(document_text):
    try:
        chunks = chunk_document(document_text)
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            responses = list(executor.map(process_chunk, chunks))
        
        combined_response = '\n'.join(responses)
        final_response = post_process_response(combined_response)
        
        return final_response
    except Exception as e:
        return f'Error in processing: {str(e)}'

def post_process_response(response):
    lines = response.split('\n')
    processed_fields = {}
    
    for line in lines:
        if ':' in line:
            key, value = line.split(':', 1)
            key = key.strip()
            value = value.strip().replace('*', '')
            
            if key not in processed_fields:
                processed_fields[key] = value
            elif value and value != processed_fields[key] and value != "Not specified":
                processed_fields[key] += f"; {value}"
    
    final_response = '\n'.join([f"{key}: {value}" for key, value in processed_fields.items()])
    
    return final_response

def extract_text_from_uploaded_file(uploaded_file):
    bytes_data = uploaded_file.read()
    doc = docx.Document(io.BytesIO(bytes_data))
    return '\n'.join([para.text for para in doc.paragraphs])

def create_formatted_docx(content, output_file):
    doc = docx.Document()
    
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    
    header_style = doc.styles.add_style('Header Style', WD_STYLE_TYPE.PARAGRAPH)
    header_style.font.name = 'Calibri'
    header_style.font.size = Pt(16)
    header_style.font.bold = True
    header_style.font.color.rgb = RGBColor(0, 0, 128)  # Navy Blue
    
    header = doc.add_paragraph("Lease Synopsis", style='Header Style')
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False
    
    table.columns[0].width = Cm(7)
    table.columns[1].width = Cm(12)
    
    lines = content.split('\n')
    for line in lines:
        if ':' in line:
            key, value = line.split(':', 1)
            row_cells = table.add_row().cells
            row_cells[0].text = key.strip()
            row_cells[1].text = value.strip().replace('*', '')
            
            key_para = row_cells[0].paragraphs[0]
            key_para.runs[0].bold = True
            key_para.runs[0].font.color.rgb = RGBColor(0, 0, 128)
            
            value_para = row_cells[1].paragraphs[0]
            value_para.runs[0].font.color.rgb = RGBColor(0, 0, 0)
    
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    doc.save(output_file)

def consolidate_synopses(all_synopses):
    consolidated = {}
    for synopsis in all_synopses:
        lines = synopsis.split('\n')
        for line in lines:
            if ':' in line:
                key, value = line.split(':', 1)
                key = key.strip()
                value = value.strip().replace('*', '')
                if key not in consolidated:
                    consolidated[key] = value
                elif value and value != consolidated[key] and value != "Not specified":
                    consolidated[key] += f"; {value}"
    
    return '\n'.join([f"{key}: {value}" for key, value in consolidated.items()])

def summarize_consolidated_synopsis(consolidated_synopsis):
    summary_prompt_template = """
    You are an expert in lease document analysis. Be concise and only return the information requested of you. Given the following concatenated lease document information, simplify the redundant parts of the text but maintain all the relevant detail:
    Consolidated Information: {consolidated_synopsis}
    """
    
    summary_prompt = PromptTemplate(template=summary_prompt_template, input_variables=["consolidated_synopsis"])
    summary_chain = LLMChain(llm=chat_llm, prompt=summary_prompt)
    
    chunks = chunk_document(consolidated_synopsis, max_tokens=8000)
    summarized_chunks = []
    
    for chunk in chunks:
        summarized_chunk = summary_chain.invoke({"consolidated_synopsis": chunk}).get('text', '').strip()
        summarized_chunks.append(summarized_chunk)
    
    return '\n\n'.join(summarized_chunks)

def create_download_zip(folder_path):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                zip_file.write(file_path, os.path.basename(file_path))
    return zip_buffer.getvalue()

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
    
    consolidated_report = consolidate_synopses(all_synopses)
    summarized_report = summarize_consolidated_synopsis(consolidated_report)
    
    create_formatted_docx(summarized_report, os.path.join(reports_folder, "consolidated_synopsis.docx"))
    
    return reports_folder

# File preview function
def show_file_preview(uploaded_file):
    if uploaded_file is not None:
        st.write("File Preview:")
        doc = docx.Document(io.BytesIO(uploaded_file.getvalue()))
        for para in doc.paragraphs[:5]:  # Show first 5 paragraphs
            st.write(para.text)
        st.write("...")

# Chatbot function
def chatbot(user_input):
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are a helpful assistant specializing in lease document analysis."},
            {"role": "user", "content": user_input}
        ]
    )
    return response.choices[0].message.content

# Streamlit UI
st.title("📄 Lease Synopsis Generator")
st.markdown("---")

uploaded_files = st.file_uploader("Upload .docx files", type="docx", accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully")
    
    # File preview
    if len(uploaded_files) == 1:
        show_file_preview(uploaded_files[0])
    else:
        selected_file = st.selectbox("Select a file to preview", uploaded_files)
        show_file_preview(selected_file)
