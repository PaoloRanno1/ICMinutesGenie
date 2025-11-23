# IC Minutes Generator

## Overview

The IC Minutes Generator is a Streamlit-based web application that automates the creation of Investment Committee (IC) minutes from PDF documents. The system uses Google's Gemini AI vision model to analyze PDF content, extract relevant information, and generate formatted IC minutes in DOCX format. The application is designed to streamline the process of converting investment committee meeting documents into professional, structured minutes.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture
- **Streamlit Framework**: Web-based interface providing file upload, progress tracking, and download capabilities
- **Single Page Application**: Streamlined user experience with real-time feedback and status updates
- **Wide Layout Configuration**: Optimized for document review and processing workflows

### Backend Architecture
- **Modular Design**: Core functionality separated into `IC_functions.py` module for reusability
- **Class-based Structure**: `ICMemoAnalyzer` class encapsulates all PDF processing and AI analysis logic
- **Stateless Processing**: Each PDF analysis is independent, allowing for concurrent usage

### Document Processing Pipeline
- **PDF to Image Conversion**: Uses PyMuPDF (fitz) to convert PDF pages to high-resolution images
- **AI Vision Analysis**: Leverages Google Gemini 2.5 Flash Lite model for content extraction and analysis
- **Document Generation**: Creates formatted DOCX outputs using python-docx library
- **Temporary File Management**: Secure handling of uploaded files with automatic cleanup

### Data Flow
1. PDF upload through Streamlit interface
2. Temporary file creation for processing
3. PDF-to-image conversion with configurable DPI
4. AI-powered content analysis and extraction
5. Structured IC minutes generation
6. DOCX document creation with professional formatting
7. Download delivery to user

### Error Handling and Validation
- **Environment Variable Validation**: Checks for required API keys on startup
- **File Type Validation**: Ensures only PDF files are processed
- **Progress Tracking**: Real-time status updates for user feedback
- **Exception Management**: Graceful error handling with user-friendly messages

## External Dependencies

### AI Services
- **Google Generative AI (Gemini)**: Primary AI model for document analysis and content extraction
- **API Key Authentication**: Requires GOOGLE_API_KEY environment variable

### Document Processing Libraries
- **PyMuPDF (fitz)**: PDF reading and image conversion capabilities
- **PIL (Pillow)**: Image processing and format handling
- **python-docx**: DOCX document generation with formatting controls

### Web Framework
- **Streamlit**: Complete web application framework with built-in UI components
- **Temporary File Handling**: Python tempfile module for secure file management

### Development Tools
- **python-dotenv**: Environment variable management for configuration
- **pathlib**: Cross-platform path handling
- **datetime**: Date and time formatting for document timestamps
- **json**: Data serialization for AI model interactions
- **re**: Regular expression processing for text formatting

### Optional Dependencies
- **IPython**: Jupyter notebook integration for development workflows (referenced in workflow scripts)
