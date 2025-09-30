import streamlit as st
import os
from server.onedrive_client import OneDriveClient
from server.document_processor import DocumentProcessor
from server.llm_service import LLMService
from utils.helpers import format_file_size, truncate_text
import asyncio
import traceback

# Initialize session state
if 'conversation_history' not in st.session_state:
    st.session_state.conversation_history = []
if 'current_document' not in st.session_state:
    st.session_state.current_document = None
if 'document_content' not in st.session_state:
    st.session_state.document_content = None
if 'original_content' not in st.session_state:
    st.session_state.original_content = None
if 'onedrive_files' not in st.session_state:
    st.session_state.onedrive_files = []

# Initialize services
@st.cache_resource
def get_services():
    return {
        'onedrive': OneDriveClient(),
        'doc_processor': DocumentProcessor(),
        'llm': LLMService()
    }

def main():
    st.title("📝 Word Document LLM Assistant")
    st.markdown("*AI-powered editing for your Microsoft 365 Word documents*")
    
    services = get_services()
    
    # Sidebar for document selection
    with st.sidebar:
        st.header("📁 OneDrive Documents")
        
        if st.button("🔄 Refresh Files", use_container_width=True):
            with st.spinner("Loading OneDrive files..."):
                try:
                    files = asyncio.run(services['onedrive'].list_word_documents())
                    st.session_state.onedrive_files = files
                    st.success(f"Found {len(files)} Word documents")
                except Exception as e:
                    st.error(f"Failed to load files: {str(e)}")
                    st.session_state.onedrive_files = []
        
        # Display files
        if st.session_state.onedrive_files:
            st.subheader("Available Documents")
            for file in st.session_state.onedrive_files:
                col1, col2 = st.columns([3, 1])
                with col1:
                    file_name = truncate_text(file['name'], 25)
                    if st.button(f"📄 {file_name}", key=f"file_{file['id']}", use_container_width=True):
                        with st.spinner("Loading document..."):
                            try:
                                content = asyncio.run(services['onedrive'].download_document(file['id']))
                                doc_text = services['doc_processor'].extract_text(content)
                                st.session_state.current_document = file
                                st.session_state.document_content = doc_text
                                st.session_state.original_content = doc_text
                                st.success(f"Loaded: {file['name']}")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Failed to load document: {str(e)}")
                
                with col2:
                    size = format_file_size(file.get('size', 0))
                    st.caption(size)
        else:
            st.info("Click 'Refresh Files' to load your Word documents from OneDrive")
    
    # Main content area
    if st.session_state.current_document:
        # Document header
        st.header(f"📄 {st.session_state.current_document['name']}")
        
        # Tabs for different views
        tab1, tab2, tab3 = st.tabs(["💬 Chat & Edit", "👀 Document Preview", "📊 Chat History"])
        
        with tab1:
            # Chat interface
            st.subheader("Chat with AI Assistant")
            
            # Display current document content in an expandable section
            with st.expander("📖 Current Document Content", expanded=False):
                if st.session_state.document_content:
                    st.text_area(
                        "Document Text",
                        value=st.session_state.document_content,
                        height=300,
                        disabled=True,
                        key="doc_preview"
                    )
            
            # Chat input
            user_message = st.chat_input("Ask me to edit your document... (e.g., 'Make it more professional', 'Add bullet points', 'Fix grammar')")
            
            if user_message:
                # Add user message to history
                from datetime import datetime
                st.session_state.conversation_history.append({
                    "role": "user",
                    "content": user_message,
                    "timestamp": datetime.now().isoformat()
                })
                
                with st.spinner("AI is processing your request..."):
                    try:
                        # Get AI response and edited content
                        response = services['llm'].process_edit_request(
                            st.session_state.document_content,
                            user_message
                        )
                        
                        # Update document content
                        st.session_state.document_content = response['edited_content']
                        
                        # Add AI response to history
                        st.session_state.conversation_history.append({
                            "role": "assistant",
                            "content": response['explanation'],
                            "timestamp": datetime.now().isoformat()
                        })
                        
                        st.success("Document updated successfully!")
                        st.rerun()
                        
                    except Exception as e:
                        error_msg = f"Failed to process request: {str(e)}"
                        st.error(error_msg)
                        st.session_state.conversation_history.append({
                            "role": "assistant",
                            "content": error_msg,
                            "timestamp": datetime.now().isoformat()
                        })
            
            # Save document button
            col1, col2, col3 = st.columns([1, 1, 1])
            with col2:
                if st.button("💾 Save to OneDrive", use_container_width=True, type="primary"):
                    if st.session_state.document_content != st.session_state.original_content:
                        with st.spinner("Saving document to OneDrive..."):
                            try:
                                # Create new document with edited content
                                new_doc_bytes = services['doc_processor'].create_document(
                                    st.session_state.document_content
                                )
                                
                                # Upload to OneDrive
                                asyncio.run(services['onedrive'].upload_document(
                                    st.session_state.current_document['id'],
                                    new_doc_bytes
                                ))
                                
                                st.session_state.original_content = st.session_state.document_content
                                st.success("✅ Document saved successfully!")
                                
                            except Exception as e:
                                st.error(f"Failed to save document: {str(e)}")
                    else:
                        st.info("No changes to save")
        
        with tab2:
            # Document preview
            st.subheader("Document Preview")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Original Content**")
                st.text_area(
                    "Original",
                    value=st.session_state.original_content or "No content loaded",
                    height=400,
                    disabled=True,
                    key="original_preview"
                )
            
            with col2:
                st.markdown("**Current Content** (with edits)")
                st.text_area(
                    "Current",
                    value=st.session_state.document_content or "No content loaded",
                    height=400,
                    disabled=True,
                    key="current_preview"
                )
            
            # Show if there are unsaved changes
            if st.session_state.document_content != st.session_state.original_content:
                st.warning("⚠️ You have unsaved changes")
        
        with tab3:
            # Chat history
            st.subheader("Conversation History")
            
            if st.session_state.conversation_history:
                for i, message in enumerate(st.session_state.conversation_history):
                    with st.chat_message(message["role"]):
                        st.write(message["content"])
                        if "timestamp" in message:
                            st.caption(f"Timestamp: {message['timestamp']}")
                
                # Clear history button
                if st.button("🗑️ Clear History"):
                    st.session_state.conversation_history = []
                    st.rerun()
            else:
                st.info("No conversation history yet. Start chatting to see messages here!")
    
    else:
        # Welcome screen
        st.markdown("""
        ## Welcome! 👋
        
        To get started:
        1. **Connect to OneDrive**: Click "Refresh Files" in the sidebar
        2. **Select a document**: Choose a Word document from your OneDrive
        3. **Start chatting**: Use natural language to edit your document
        
        ### What can I help you with?
        - Fix grammar and spelling
        - Improve writing style and tone
        - Add or remove content
        - Format and structure documents
        - Summarize or expand sections
        - And much more!
        """)
        
        # Sample commands
        with st.expander("💡 Example Commands"):
            st.markdown("""
            - *"Make this more professional"*
            - *"Add bullet points to the main ideas"*
            - *"Fix any grammar mistakes"*
            - *"Make it shorter and more concise"*
            - *"Add a conclusion paragraph"*
            - *"Change the tone to be more casual"*
            """)

if __name__ == "__main__":
    # Check for required environment variables
    if not os.getenv("OPENAI_API_KEY"):
        st.error("⚠️ OPENAI_API_KEY environment variable is required")
        st.stop()
    
    main()
