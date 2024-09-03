
import streamlit as st
import os
import tempfile
import shutil
from improved_lease_synopsis_generator import process_docx_files, extract_text_from_docx
from langchain.embeddings import OpenAIEmbeddings
from langchain.vectorstores import FAISS  # Assuming FAISS is the most suitable for your needs
from langchain.chat_models import ChatOpenAI
from langchain.chains import ConversationalRetrievalChain
from langchain.memory import ConversationBufferMemory

st.set_page_config(layout='wide', page_title="Lease Synopsis Generator", page_icon="📄")

# Custom CSS for consistent styling
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
    .document-preview {
        border: 1px solid #ddd;
        padding: 10px;
        margin-bottom: 10px;
        border-radius: 5px;
    }
    .document-preview h3 {
        margin-top: 0;
        color: #2C3E50;
    }
</style>
""", unsafe_allow_html=True)

# App Title and Intro
st.title("📄 Lease Synopsis Generator and Chatbot")
st.markdown("---")

# Initialize session state variables
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'vector_store' not in st.session_state:
    st.session_state.vector_store = None
if 'document_previews' not in st.session_state:
    st.session_state.document_previews = {}

# File Upload Section
uploaded_files = st.file_uploader("Upload .docx files", type="docx", accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully")
    
    if st.button("Generate Lease Synopsis and Prepare Chatbot"):
        with st.spinner("🔎 Generating lease synopsis and preparing chatbot..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save uploaded files and extract text
                documents = []
                for uploaded_file in uploaded_files:
                    file_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    text = extract_text_from_docx(file_path)
                    documents.append(text)
                    
                    # Create document preview
                    preview = text[:500] + "..." if len(text) > 500 else text
                    st.session_state.document_previews[uploaded_file.name] = preview
                
                # Generate Lease Synopses
                reports_folder = process_docx_files(temp_dir)
                
                # Zip the synopses for download
                zip_path = shutil.make_archive(os.path.join(temp_dir, "lease_synopses"), 'zip', reports_folder)
                
                st.success("✅ Lease synopses generated successfully")
                
                # Download button for the generated synopses
                with open(zip_path, "rb") as file:
                    btn = st.download_button(
                        label="Download Lease Synopses",
                        data=file,
                        file_name="lease_synopses.zip",
                        mime="application/zip"
                    )
                
                # Create the vector store using FAISS
                embeddings = OpenAIEmbeddings()
                st.session_state.vector_store = FAISS.from_texts(documents, embeddings)
                
                st.success("✅ Chatbot prepared successfully")

    # Display Document Previews
    if st.session_state.document_previews:
        st.subheader("Document Previews")
        for filename, preview in st.session_state.document_previews.items():
            with st.expander(f"Preview: {filename}"):
                st.markdown(f"<div class='document-preview'><h3>{filename}</h3>{preview}</div>", unsafe_allow_html=True)

    # Chatbot Interface
    if st.session_state.vector_store is not None:
        st.subheader("Chat with your Lease Documents")
        
        # Initialize the Conversational Chain
        llm = ChatOpenAI(temperature=0)
        memory = ConversationBufferMemory(memory_key="chat_history", return_messages=True)
        conversation_chain = ConversationalRetrievalChain.from_llm(
            llm=llm,
            retriever=st.session_state.vector_store.as_retriever(),
            memory=memory
        )
        
        # User Input for Chatbot
        user_question = st.text_input("Ask a question about your lease documents:")
        if user_question:
            response = conversation_chain({"question": user_question})
            st.session_state.chat_history.append(("You", user_question))
            st.session_state.chat_history.append(("AI", response['answer']))
        
        # Display the chat history
        for role, message in st.session_state.chat_history:
            if role == "You":
                st.write(f"👤 **You:** {message}")
            else:
                st.write(f"🤖 **AI:** {message}")

else:
    st.warning("⚠️ Please upload .docx files to generate lease synopses and prepare the chatbot")

st.markdown("<br>" * 15, unsafe_allow_html=True)
st.markdown("---")
st.markdown("Created with ❤️ by Weaver")
