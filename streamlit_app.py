# Generate Lease Synopses
        reports_folder = process_docx_files(temp_dir)
        
        # Zip the synopses for download
        zip_path = shutil.make_archive(os.path.join(temp_dir, "lease_synopses"), 'zip', reports_folder)
        
        st.success("✅ Lease synopses generated successfully")
        
        # Download button for the generated synopses
        with open(zip_path, "rb") as file:
            st.download_button(
                label="Download Lease Synopses",
                data=file,
                file_name="lease_synopses.zip",
                mime="application/zip"
            )
        
        # Create the vector store using FAISS
        embeddings = OpenAIEmbeddings()
        st.session_state.vector_store = FAISS.from_texts(documents, embeddings)
        
        st.success("✅ Chatbot prepared successfully")

def handle_user_input(user_question: str, conversation_chain) -> None:
    """Process user input and generate AI response."""
    try:
        # Split the user question if it's too long
        question_chunks = chunk_document(user_question, max_tokens=4000)
        responses = []
        
        for chunk in question_chunks:
            response = conversation_chain({"question": chunk})
            responses.append(response['answer'])
        
        combined_response = ' '.join(responses)
        
        st.session_state.chat_history.append(("You", user_question))
        st.session_state.chat_history.append(("AI", combined_response))
    except (APIError, RateLimitError, APIConnectionError, APITimeoutError) as e:
        st.error(f"An error occurred with the OpenAI API: {str(e)}")
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
        st.error(traceback.format_exc())

def show_file_preview(uploaded_file):
    """Display a preview of the uploaded file."""
    if uploaded_file is not None:
        st.write("File Preview:")
        doc = docx.Document(uploaded_file)
        for para in doc.paragraphs[:5]:  # Show first 5 paragraphs
            st.write(para.text)
        st.write("...")

def main():
    # File Upload Section
    uploaded_files = st.file_uploader("Upload .docx files", type="docx", accept_multiple_files=True)

    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully")
        
        # File preview
        if len(uploaded_files) == 1:
            show_file_preview(uploaded_files[0])
        else:
            selected_file = st.selectbox("Select a file to preview", uploaded_files)
            show_file_preview(selected_file)
        
        if st.button("Generate Lease Synopsis and Prepare Chatbot"):
            with st.spinner("🔎 Generating lease synopsis and preparing chatbot..."):
                try:
                    process_uploaded_files(uploaded_files)
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    st.error("Stack trace:", exc_info=True)
        
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
                handle_user_input(user_question, conversation_chain)
            
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

if __name__ == "__main__":
    main()
