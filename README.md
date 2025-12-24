# PDF Text Extractor

A modern, feature-rich PDF text extraction tool with OCR capabilities built with PyQt5 and Tesseract OCR.

![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)

## Features

- **Fast Text Extraction** - Extract text from PDF files using pdfplumber
- **OCR Support** - Advanced OCR capabilities with Tesseract for scanned documents
- **Multi-Language** - Support for multiple languages via Tesseract
- **Fast Mode** - Toggle between high-quality (300 DPI) and fast (200 DPI) OCR processing
- **Multiple Export Formats** - Save extracted text as TXT or DOCX
- **Modern Dark UI** - Beautiful dark-themed interface with smooth animations
- **Progress Tracking** - Real-time progress bar and status updates
- **Cancellation Support** - Cancel long-running extractions at any time
- **Smart OCR** - Optionally use OCR only on pages without native text

## Installation

### Prerequisites

- Python 3.7 or higher
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (optional, for OCR functionality)

### Install Tesseract OCR

#### Windows
1. Download the installer from [Tesseract at UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
2. Run the installer and follow the instructions
3. The application will automatically detect Tesseract

#### Linux (Ubuntu/Debian)
```bash
sudo apt-get update
sudo apt-get install tesseract-ocr
sudo apt-get install tesseract-ocr-eng  # English language pack
```

#### macOS
```bash
brew install tesseract
```

### Install Python Dependencies

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/PDF_File_Extractor.git
cd PDF_File_Extractor

# Install required packages
pip install -r requirements.txt
```

## Usage

### Running the Application

```bash
python "PDF-Reader.py"
```

### Basic Workflow

1. **Select PDF** - Click "Select PDF" to choose a PDF file
2. **Configure OCR** (if needed):
   - Select OCR language from the dropdown
   - Toggle "Fast Mode" for quicker processing
   - Choose whether to use OCR only on pages without text
3. **Extract** - Click "Extract Text" to start extraction
4. **Save** - Save the extracted text as TXT or DOCX

## Requirements

```
PyQt5>=5.15.0
pdfplumber>=0.10.0
pytesseract>=0.3.10
Pillow>=10.0.0
python-docx>=0.8.11
```

## Key Features Explained

### Fast Mode
- **Enabled**: Uses 200 DPI resolution and minimal preprocessing (faster)
- **Disabled**: Uses 300 DPI resolution with median filtering (higher quality)

### OCR Language Support
The application automatically detects available Tesseract language packs. Install additional languages:

```bash
# Windows - Install during Tesseract setup
# Linux
sudo apt-get install tesseract-ocr-ita  # Italian
sudo apt-get install tesseract-ocr-fra  # French
sudo apt-get install tesseract-ocr-deu  # German
```

### Smart OCR Processing
When "Use OCR only on pages without text" is enabled:
- Pages with native text are extracted directly (fast)
- Only scanned/image-based pages use OCR (slower but necessary)

## Configuration

The application creates the following files:
- `pdf_extractor.log` - Application logs
- `assets/app_icon.png` - Auto-generated application icon
- `pyrightconfig.json` - Python type checking configuration

## Troubleshooting

### Tesseract Not Found
If you see "Tesseract OCR was not found":
1. Install Tesseract OCR (see Installation section)
2. Restart the application
3. The app will work for native PDF text extraction even without Tesseract

### Language Not Available
If your desired OCR language isn't listed:
1. Install the language pack for Tesseract
2. Restart the application
3. The language should appear in the dropdown

### Slow OCR Processing
For faster OCR:
1. Enable "Fast Mode" checkbox
2. Use "OCR only on pages without text" option
3. Consider using a machine with better CPU

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [pdfplumber](https://github.com/jsvine/pdfplumber) - PDF text extraction
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) - OCR engine
- [PyQt5](https://www.riverbankcomputing.com/software/pyqt/) - GUI framework
- [Pillow](https://python-pillow.org/) - Image processing

## Contact

For questions or support, please open an issue on GitHub.

---

Made with Python and PyQt5
