import streamlit as st
import os
import tempfile
from pathlib import Path
import traceback
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Import the IC functions module
import IC_functions

# Set page configuration
st.set_page_config(
    page_title="IC Minutes Generator",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

def initialize_analyzer():
    """Initialize the IC Memo Analyzer with API key from environment"""
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        st.error("⚠️ Google API Key not found. Please set the GOOGLE_API_KEY environment variable.")
        st.stop()
    
    try:
        analyzer = IC_functions.ICMemoAnalyzer(api_key, model_name="gemini-2.5-flash-lite")
        return analyzer
    except Exception as e:
        st.error(f"❌ Failed to initialize IC Memo Analyzer: {str(e)}")
        st.stop()

def process_pdf(analyzer, uploaded_file):
    """Process the uploaded PDF and generate IC minutes"""
    
    # Create a temporary file to save the uploaded PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_file_path = tmp_file.name
    
    try:
        # Create progress indicators
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Step 1: Analyze PDF
        status_text.text("📄 Converting PDF to images and analyzing content...")
        progress_bar.progress(25)
        
        results = analyzer.analyze_pdf(tmp_file_path, delay=1.0)
        
        if not results:
            st.error("❌ No content could be extracted from the PDF.")
            return None, None, None
        
        # Check if analysis was successful
        successful_pages = [r for r in results if r.get("status") == "success"]
        if not successful_pages:
            st.error("❌ Failed to analyze any pages in the PDF.")
            return None, None, None
        
        progress_bar.progress(60)
        status_text.text("🤖 Generating IC Minutes...")
        
        # Step 2: Generate IC Minutes
        ic_minutes = analyzer.create_IC_minutes(results)
        
        if not ic_minutes:
            st.error("❌ Failed to generate IC Minutes.")
            return None, None, None
        
        progress_bar.progress(80)
        status_text.text("📝 Creating Word document...")
        
        # Step 3: Generate .docx file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as docx_file:
            docx_path = docx_file.name
        
        analyzer.output_docx(ic_minutes, path=docx_path)
        
        progress_bar.progress(100)
        status_text.text("✅ IC Minutes generated successfully!")
        
        return ic_minutes, docx_path, results
        
    except Exception as e:
        st.error(f"❌ Error processing PDF: {str(e)}")
        st.error("Please check the PDF format and try again.")
        return None, None, None
    
    finally:
        # Clean up temporary PDF file
        try:
            os.unlink(tmp_file_path)
        except:
            pass

def check_password():
    """Returns True if the user has entered the correct password."""
    
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == "StradaLegal2025":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show password input
        st.title("🔐 Strada Partners Access")
        st.markdown("**Please enter the access password to continue**")
        st.text_input(
            "Password", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.markdown("---")
        st.markdown(
            "<div style='text-align: center; color: #666; margin-top: 50px;'>"
            "Authorized personnel only | Strada Partners"
            "</div>",
            unsafe_allow_html=True
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password incorrect
        st.title("🔐 Strada Partners Access")
        st.markdown("**Please enter the access password to continue**")
        st.text_input(
            "Password", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("❌ Incorrect password. Please try again.")
        st.markdown("---")
        st.markdown(
            "<div style='text-align: center; color: #666; margin-top: 50px;'>"
            "Authorized personnel only | Strada Partners"
            "</div>",
            unsafe_allow_html=True
        )
        return False
    else:
        # Password correct
        return True

def main():
    """Main Streamlit application"""
    
    # Check password before showing main app
    if not check_password():
        return
    
    # Header
    st.title("📄 IC Minutes Generator")
    st.markdown("**Convert Investment Committee memos into formatted IC Minutes**")
    st.markdown("---")
    
    # Initialize the analyzer
    analyzer = initialize_analyzer()
    
    # Sidebar with information
    with st.sidebar:
        st.header("ℹ️ About")
        st.markdown("""
        This application converts IC (Investment Committee) memo PDFs into properly formatted IC Minutes documents.
        
        **How it works:**
        1. Upload your IC memo PDF
        2. AI analyzes the document content  
        3. Generates structured IC Minutes
        4. Download as formatted .docx file
        
        **Supported formats:**
        - PDF files only
        - Multi-page documents supported
        """)
        
        st.header("🔧 Processing Details")
        st.markdown("""
        - Uses Google Gemini AI for content analysis
        - Extracts key information from memo pages
        - Generates structured output with:
          - Context
          - Investment Overview  
          - Ask
          - Conclusion
        """)
    
    # Main content area
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("📁 Upload PDF")
        uploaded_file = st.file_uploader(
            "Choose your IC memo PDF file",
            type=['pdf'],
            help="Upload a PDF file containing the Investment Committee memo to be processed."
        )
        
        if uploaded_file is not None:
            st.success(f"✅ File uploaded: {uploaded_file.name}")
            st.info(f"📊 File size: {len(uploaded_file.getvalue()) / 1024:.1f} KB")
            
            # Process button
            if st.button("🚀 Generate IC Minutes", type="primary"):
                with st.spinner("Processing your document..."):
                    ic_minutes, docx_path, analysis_results = process_pdf(analyzer, uploaded_file)
                    
                    if ic_minutes and docx_path:
                        # Store results in session state
                        st.session_state.ic_minutes = ic_minutes
                        st.session_state.docx_path = docx_path
                        st.session_state.original_filename = uploaded_file.name
                        st.session_state.analysis_results = analysis_results
                        st.rerun()
    
    with col2:
        st.subheader("📋 Generated IC Minutes")
        
        # Display results if available
        if hasattr(st.session_state, 'ic_minutes') and st.session_state.ic_minutes:
            
            # Download button
            try:
                with open(st.session_state.docx_path, "rb") as file:
                    docx_data = file.read()
                
                # Create download filename
                original_name = Path(st.session_state.original_filename).stem
                download_filename = f"IC_Minutes_{original_name}.docx"
                
                st.download_button(
                    label="📥 Download IC Minutes (.docx)",
                    data=docx_data,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )
                
            except Exception as e:
                st.error(f"❌ Error preparing download: {str(e)}")
            
            # Display the content
            st.markdown("### Preview:")
            
            # Format the IC minutes for better display
            formatted_content = st.session_state.ic_minutes.replace('\n\n', '\n\n---\n\n')
            
            with st.expander("📖 View Full Content", expanded=True):
                st.markdown(formatted_content)
                
            # Show word count
            word_count = len(st.session_state.ic_minutes.split())
            st.caption(f"📊 Word count: {word_count} words")
            
            # Add button to show raw analysis results
            st.markdown("---")
            if st.button("🔍 Show Raw Analysis Results", help="View detailed page-by-page analysis from AI"):
                if hasattr(st.session_state, 'analysis_results') and st.session_state.analysis_results:
                    with st.expander("📄 Raw Analysis Results", expanded=True):
                        for result in st.session_state.analysis_results:
                            page_num = result.get('page_number', 'Unknown')
                            status = result.get('status', 'unknown')
                            content = result.get('content', 'No content')
                            
                            if status == 'success':
                                st.markdown(f"### Page {page_num}")
                                st.markdown(content)
                                st.markdown("---")
                            else:
                                st.error(f"**Page {page_num}**: {content}")
                else:
                    st.warning("No analysis results available. Please process a PDF first.")
            
        else:
            st.info("👆 Upload a PDF file and click 'Generate IC Minutes' to see results here.")
            
            # Show example of what the output looks like
            with st.expander("💡 Example Output Structure"):
                st.markdown("""
                **The generated IC Minutes will include:**
                
                **1. Context**
                - Company introduction and background
                - Business model overview
                - Industry positioning
                
                **2. Investment Overview**  
                - Detailed investment description
                - Financial performance analysis
                - Strategic rationale and growth drivers
                - Risk assessment and mitigation
                
                **3. Ask**
                - Specific approval sought from Investment Committee
                
                **4. Conclusion**
                - Formal approval statement
                """)

    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666;'>"
        "Powered by Google Gemini AI | Strada Partners IC Minutes Generator"
        "</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
