==================================================
PDF TEXT EXTRACTOR
==================================================

A powerful Python application for extracting text from PDF files with OCR capabilities for scanned documents.

PROJECT: PDF Text Extractor
VERSION: 1.0
LANGUAGE: Python 3.6+
GUI: PyQt5

==================================================
FEATURES
==================================================

- Extract text from both digital and scanned PDFs
- Automatic OCR (Optical Character Recognition) for scanned pages
- Multi-language support (English, Italian, French, German, Spanish, etc.)
- Modern and intuitive graphical interface
- Real-time progress tracking
- Save extracted text as TXT or DOCX files
- Error logging and handling
- Support for high-resolution PDF processing

==================================================
SYSTEM REQUIREMENTS
==================================================

OPERATING SYSTEM:
- Windows 7/8/10/11
- macOS 10.12+
- Linux (Ubuntu 16.04+, CentOS 7+, etc.)

SOFTWARE DEPENDENCIES:
- Python 3.6 or higher
- Tesseract OCR 4.0+

==================================================
INSTALLATION GUIDE
==================================================

STEP 1: Install Python Dependencies

Create and activate a virtual environment (recommended):
```bash
python -m venv pdf_env
source pdf_env/bin/activate  # Linux/macOS
pdf_env\Scripts\activate     # Windows

STEP 2: Installation of packages

Type in the terminal the command:
pip install -r requirements.txt

Manual installation of packages:
pip install PyQt5 python-docx pdfplumber pytesseract Pillow
