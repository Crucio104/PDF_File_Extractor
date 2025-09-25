================================================================================
                           PDF Text Extractor v2.0
================================================================================

A modern, user-friendly GUI application for extracting text from PDF files 
with advanced OCR capabilities.

================================================================================
FEATURES
================================================================================

* Dual Text Extraction: Combines direct text extraction with OCR for scanned documents
* Smart OCR Processing: Option to use OCR only on pages without extractable text
* Multi-language OCR Support: Supports all Tesseract language packages
* Multiple Output Formats: Save as TXT or DOCX files
* Modern UI: Clean, responsive interface with custom styling
* Progress Tracking: Real-time progress bar and status updates
* Cross-platform: Works on Windows, macOS, and Linux
* Automatic Tesseract Detection: Finds Tesseract installation automatically
* Robust Error Handling: Comprehensive error messages and logging

================================================================================
PREREQUISITES
================================================================================

REQUIRED SOFTWARE
-----------------

1. Python 3.8 or higher
   Download from: https://www.python.org/downloads/

2. Tesseract OCR
   Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki
   macOS: brew install tesseract
   Ubuntu/Debian: sudo apt install tesseract-ocr
   CentOS/RHEL: sudo yum install tesseract

OPTIONAL LANGUAGE PACKS
-----------------------

For OCR in languages other than English:
Windows: Download language data files from Tesseract GitHub releases
macOS: brew install tesseract-lang
Linux: sudo apt install tesseract-ocr-[language-code]

Example for Italian: tesseract-ocr-ita

================================================================================
INSTALLATION
================================================================================

METHOD 1: USING PIP (RECOMMENDED)
----------------------------------

1. Clone or download the project:
   git clone <repository-url>
   cd pdf-text-extractor

2. Create a virtual environment (recommended):
   python -m venv venv
   
   On Windows:
   venv\Scripts\activate
   
   On macOS/Linux:
   source venv/bin/activate

3. Install dependencies:
   pip install -r requirements.txt

METHOD 2: MANUAL INSTALLATION
-----------------------------

pip install pdfplumber>=0.10.0
pip install pytesseract>=0.3.10
pip install Pillow>=10.0.0
pip install python-docx>=0.8.11
pip install PyQt5>=5.15.9

================================================================================
USAGE
================================================================================

RUNNING THE APPLICATION
-----------------------

python pdf_extractor.py

BASIC WORKFLOW
--------------

1. Launch the application
2. Select a PDF file using the "Select PDF" button
3. Choose OCR language from the dropdown (if available)
4. Configure OCR settings using the checkbox option
5. Click "Extract Text" to start the process
6. Review the extracted text in the text area
7. Save the result as TXT or DOCX format

OCR OPTIONS
-----------

* "Use OCR only on pages without text" (default): Uses OCR only when 
  direct text extraction fails
* Unchecked: Forces OCR processing on all pages (slower but may capture 
  more content)

================================================================================
FILE STRUCTURE
================================================================================

pdf-text-extractor/
├── pdf_extractor.py      # Main application file
├── requirements.txt      # Python dependencies
├── README.txt           # This file
├── pdf_extractor.log    # Application logs (created at runtime)
└── assets/             # Optional icon directory
    └── pdf.ico         # Application icon (optional)

================================================================================
TROUBLESHOOTING
================================================================================

COMMON ISSUES
-------------

1. "Tesseract not found" error
   - Ensure Tesseract is installed and in your system PATH
   - On Windows, the application tries common installation paths automatically
   - For custom installations, verify the path in system environment variables

2. "Language file not found" error
   - Install the required language pack for Tesseract
   - Verify the language code is correct (e.g., 'eng' for English, 'ita' for Italian)

3. Poor OCR quality
   - Try different OCR languages
   - Ensure the PDF has good image quality
   - Consider preprocessing the PDF externally for better results

4. Application won't start
   - Verify all dependencies are installed: pip list
   - Check Python version: python --version (requires 3.8+)
   - Run in debug mode to see detailed error messages

5. Memory issues with large PDFs
   - Close other applications to free memory
   - Process smaller sections of the PDF separately
   - Consider using a machine with more RAM for very large documents

LOG FILES
---------

The application creates detailed logs in pdf_extractor.log. Check this file for:
- Extraction errors
- OCR processing issues
- File I/O problems
- Performance metrics

PERFORMANCE TIPS
----------------

* For better speed: Use "OCR only on pages without text" option
* For better accuracy: Uncheck the option above and process all pages with OCR
* For large files: Monitor system resources and process during off-peak hours

================================================================================
TECHNICAL DETAILS
================================================================================

ARCHITECTURE
------------

* GUI Framework: PyQt5 for cross-platform desktop interface
* PDF Processing: pdfplumber for text extraction
* OCR Engine: Tesseract via pytesseract wrapper
* Image Processing: Pillow (PIL) for OCR preprocessing
* Document Output: python-docx for Word document creation

IMAGE PREPROCESSING
-------------------

The application applies several preprocessing steps to improve OCR accuracy:
- Conversion to grayscale
- Automatic contrast enhancement
- Binary thresholding
- Noise reduction with median filtering

THREADING
---------

* Uses QThread for non-blocking text extraction
* Progress updates in real-time
* Graceful cancellation support
* Memory management for large documents

================================================================================
SYSTEM REQUIREMENTS
================================================================================

MINIMUM REQUIREMENTS
--------------------
OS: Windows 10, macOS 10.14, or Linux with X11
Python: 3.8+
RAM: 4GB (8GB recommended for large PDFs)
Storage: 100MB free space
Display: 1024x768 minimum resolution

RECOMMENDED REQUIREMENTS
------------------------
RAM: 8GB or more
CPU: Multi-core processor for faster OCR
Display: 1920x1080 or higher
SSD: For faster file I/O operations

================================================================================
VERSION HISTORY
================================================================================

v2.0 (Current)
- Automatic Tesseract path detection
- Improved error handling and user feedback
- Modern UI with custom styling
- Multi-format save options (TXT/DOCX)
- Smart OCR processing options
- Progress tracking and status updates
- Cross-platform compatibility improvements

v1.0
- Basic PDF text extraction
- Simple OCR integration
- Basic GUI interface

================================================================================
ACKNOWLEDGMENTS
================================================================================

* pdfplumber: PDF processing library
* Tesseract: Open source OCR engine
* PyQt5: GUI framework
* Pillow: Python imaging library

================================================================================

For the most up-to-date information, please visit the project repository.

================================================================================
