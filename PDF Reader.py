import sys
import os
import logging
from pathlib import Path
from typing import Optional, List
from docx import Document
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QFileDialog, 
                             QLabel, QProgressBar, QComboBox, QMessageBox, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon, QFont
import pdfplumber
import pytesseract
from PIL import Image, ImageFilter, ImageOps

# Logging configuration
logging.basicConfig(
    filename="pdf_extractor.log", 
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    filemode='a'
)

class TesseractConfig:
    """Class to manage Tesseract configuration"""
    
    @staticmethod
    def find_tesseract_path() -> Optional[str]:
        """Automatically find Tesseract path"""
        possible_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            "/usr/bin/tesseract",
            "/usr/local/bin/tesseract",
            "/opt/homebrew/bin/tesseract"  # macOS with Homebrew
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        # Try to find tesseract in PATH
        import shutil
        tesseract_cmd = shutil.which("tesseract")
        if tesseract_cmd:
            return tesseract_cmd
            
        return None
    
    @staticmethod
    def get_available_languages(tesseract_path: str) -> List[str]:
        """Get available languages for OCR"""
        try:
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            langs = pytesseract.get_languages()
            return sorted(langs) if langs else ["eng"]
        except Exception as e:
            logging.error(f"Error retrieving languages: {e}")
            return ["eng"]

class ExtractThread(QThread):
    """Thread for text extraction in background"""
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)
    status_update = pyqtSignal(str)

    def __init__(self, file_path: str, ocr_lang: str, tesseract_path: str, 
                 ocr_only_blank: bool = True): 
        super().__init__()
        self.file_path = file_path
        self.ocr_lang = ocr_lang
        self.tesseract_path = tesseract_path
        self.ocr_only_blank = ocr_only_blank
        self._is_running = True
    
    def stop(self):
        """Stop the thread"""
        self._is_running = False
        self.quit()
        self.wait()
    
    def run(self):
        try:
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
            
            with pdfplumber.open(self.file_path) as pdf:
                extracted_text = ""
                total_pages = len(pdf.pages)
                
                self.status_update.emit(f"Processing {total_pages} pages...")
                
                for i, page in enumerate(pdf.pages):
                    if not self._is_running:
                        return
                    
                    # Attempt direct text extraction
                    text = page.extract_text()
                    has_text = text and text.strip()
                    
                    if has_text and self.ocr_only_blank:
                        extracted_text += text + "\n\n"
                        self.status_update.emit(f"Page {i+1}: text extracted directly")
                    else:
                        # OCR needed
                        self.status_update.emit(f"Page {i+1}: performing OCR...")
                        try:
                            ocr_text = self._perform_ocr(page)
                            extracted_text += ocr_text + "\n\n"
                        except Exception as e:
                            logging.error(f"OCR error on page {i+1}: {e}")
                            if has_text:  # Fallback to directly extracted text
                                extracted_text += text + "\n\n"
                            else:
                                extracted_text += f"[Error processing page {i+1}]\n\n"
                    
                    progress = int((i + 1) / total_pages * 100)
                    self.progress.emit(progress)
                
            self.finished.emit(extracted_text)
            
        except FileNotFoundError:
            self.error.emit("PDF file not found.")
        except Exception as e:
            logging.error(f"Error during extraction: {e}")
            self.error.emit(f"Error during extraction: {str(e)}")
    
    def _perform_ocr(self, page) -> str:
        """Perform OCR on a page"""
        try:
            # Convert page to high-resolution image
            pil_image = page.to_image(resolution=300).original
            
            # Image preprocessing to improve OCR
            pil_image = self._preprocess_image(pil_image)
            
            # Perform OCR
            ocr_text = pytesseract.image_to_string(pil_image, lang=self.ocr_lang)
            
            # Free memory
            pil_image.close()
            
            return ocr_text
            
        except pytesseract.TesseractNotFoundError:
            raise Exception("Tesseract not found. Please check installation.")
        except Exception as e:
            if "Failed loading language" in str(e):
                raise Exception(f"Language file '{self.ocr_lang}' not found.")
            raise e
    
    def _preprocess_image(self, image: Image.Image) -> Image.Image:
        """Preprocess image to improve OCR quality"""
        # Convert to grayscale
        image = image.convert("L")
        
        # Improve contrast
        image = ImageOps.autocontrast(image)
        
        # Binarization
        image = image.point(lambda x: 0 if x < 150 else 255, '1')
        
        # Median filter to reduce noise
        image = image.filter(ImageFilter.MedianFilter(size=3))
        
        return image

class PDFExtractorGUI(QMainWindow): 
    """Main GUI for PDF extractor"""
    
    def __init__(self):
        super().__init__()
        self.file_path: Optional[str] = None  
        self.extract_thread: Optional[ExtractThread] = None
        self.tesseract_path: Optional[str] = None
        
        self._init_tesseract()
        self.initUI()
        self._setup_styles()
        self._set_icon()
        
        # Timer to hide status messages
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self._clear_status_message)
    
    def _init_tesseract(self):
        """Initialize Tesseract"""
        self.tesseract_path = TesseractConfig.find_tesseract_path()
        if not self.tesseract_path:
            QMessageBox.warning(
                None, 
                "Tesseract not found", 
                "Tesseract OCR was not found. Please install it to use OCR functionality."
            )
    
    def _set_icon(self):
        """Set window icon"""
        # Search for icon in different locations
        icon_paths = [
            "pdf.ico",
            "assets/pdf.ico",
            os.path.join(os.path.dirname(__file__), "pdf.ico")
        ]
        
        for icon_path in icon_paths:
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
                return
        
        # Default icon if not found
        self.setWindowIcon(self.style().standardIcon(self.style().SP_FileDialogDetailedView))

    def initUI(self):
        """Initialize user interface"""
        self.setWindowTitle("PDF Text Extractor v2.0")
        self.setGeometry(100, 100, 700, 700)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        # File selection section
        self._create_file_selection_section(layout)
        
        # OCR configuration section
        self._create_ocr_config_section(layout)
        
        # Progress bar
        self._create_progress_section(layout)
        
        # Extract button
        self.extract_btn = QPushButton("Extract Text")
        self.extract_btn.clicked.connect(self.extract_text)
        self.extract_btn.setEnabled(False)
        layout.addWidget(self.extract_btn)
        
        # Text area
        self._create_text_area_section(layout)
        
        # Save section
        self._create_save_section(layout)
        
        # Status bar
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)
    
    def _create_file_selection_section(self, layout: QVBoxLayout):
        """Create file selection section"""
        file_layout = QHBoxLayout()
        
        self.file_label = QLabel("No file selected")
        self.select_btn = QPushButton("Select PDF")
        self.select_btn.clicked.connect(self.select_file)
        
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.select_btn)
        layout.addLayout(file_layout)
    
    def _create_ocr_config_section(self, layout: QVBoxLayout):
        """Create OCR configuration section"""
        ocr_layout = QVBoxLayout()
        
        # Language selection
        lang_layout = QHBoxLayout()
        lang_label = QLabel("OCR Language:")
        self.lang_box = QComboBox()
        
        if self.tesseract_path:
            languages = TesseractConfig.get_available_languages(self.tesseract_path)
            self.lang_box.addItems(languages)
            if "eng" in languages:
                self.lang_box.setCurrentText("eng")
        else:
            self.lang_box.addItem("OCR not available")
            self.lang_box.setEnabled(False)
        
        lang_layout.addWidget(lang_label)
        lang_layout.addWidget(self.lang_box)
        
        # Checkbox for OCR only on blank pages
        self.ocr_blank_only = QCheckBox("Use OCR only on pages without text")
        self.ocr_blank_only.setChecked(True)
        self.ocr_blank_only.setToolTip("If unchecked, will use OCR on all pages")
        
        ocr_layout.addLayout(lang_layout)
        ocr_layout.addWidget(self.ocr_blank_only)
        layout.addLayout(ocr_layout)
    
    def _create_progress_section(self, layout: QVBoxLayout):
        """Create progress bar section"""
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setMaximumHeight(20)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
    
    def _create_text_area_section(self, layout: QVBoxLayout):
        """Create text area section for displaying results"""
        text_label = QLabel("Extracted text:")
        layout.addWidget(text_label)
        
        self.text_area = QTextEdit()
        self.text_area.setPlaceholderText("Extracted text will appear here...")
        self.text_area.setReadOnly(True)
        
        # Set monospace font for better readability
        font = QFont("Consolas", 10)
        if not font.exactMatch():
            font = QFont("Courier New", 10)
        self.text_area.setFont(font)
        
        layout.addWidget(self.text_area)
    
    def _create_save_section(self, layout: QVBoxLayout):
        """Create save section"""
        save_layout = QHBoxLayout()
        
        self.save_btn = QPushButton("Save as TXT")
        self.save_btn.clicked.connect(lambda: self.save_text("txt"))
        self.save_btn.setEnabled(False)
        
        self.save_docx_btn = QPushButton("Save as DOCX")
        self.save_docx_btn.clicked.connect(lambda: self.save_text("docx"))
        self.save_docx_btn.setEnabled(False)
        
        save_layout.addWidget(self.save_btn)
        save_layout.addWidget(self.save_docx_btn)
        layout.addLayout(save_layout)
    
    def _setup_styles(self):
        """Set CSS styles"""
        self.setStyleSheet("""
        QMainWindow {
            background-color: #f8f9fa;
        }
        QWidget {
            background-color: #f8f9fa;
            font-family: 'Segoe UI', 'Arial', sans-serif;
            font-size: 14px;
        }
        QLabel {
            color: #2c3e50;
            font-size: 14px;
            padding: 5px;
            font-weight: 500;
        }
        QPushButton {
            background-color: #007bff;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 14px;
            min-height: 20px;
        }
        QPushButton:hover {
            background-color: #0056b3;
        }
        QPushButton:pressed {
            background-color: #004085;
        }
        QPushButton:disabled {
            background-color: #6c757d;
            color: #adb5bd;
        }
        QTextEdit {
            background-color: white;
            border: 2px solid #dee2e6;
            border-radius: 6px;
            padding: 10px;
            font-size: 13px;
            color: #2c3e50;
            line-height: 1.5;
        }
        QTextEdit:disabled {
            background-color: #f8f9fa;
            color: #6c757d;
        }
        QProgressBar {
            border: 2px solid #dee2e6;
            border-radius: 6px;
            text-align: center;
            background-color: white;
            color: #2c3e50;
            font-weight: 600;
        }
        QProgressBar::chunk {
            background-color: #28a745;
            border-radius: 4px;
        }
        QComboBox {
            background-color: white;
            border: 2px solid #dee2e6;
            border-radius: 6px;
            padding: 8px 12px;
            font-size: 14px;
            color: #2c3e50;
            min-height: 20px;
        }
        QComboBox:focus {
            border: 2px solid #007bff;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 30px;
            border-left: 2px solid #dee2e6;
            border-top-right-radius: 6px;
            border-bottom-right-radius: 6px;
            background: #f8f9fa;
        }
        QComboBox QAbstractItemView {
            background: white;
            border: 2px solid #dee2e6;
            selection-background-color: #007bff;
            selection-color: white;
            font-size: 14px;
            outline: none;
        }
        QCheckBox {
            color: #2c3e50;
            font-size: 14px;
            spacing: 8px;
        }
        QCheckBox::indicator {
            width: 18px;
            height: 18px;
            border: 2px solid #dee2e6;
            border-radius: 3px;
            background-color: white;
        }
        QCheckBox::indicator:checked {
            background-color: #007bff;
            border-color: #007bff;
        }
        QCheckBox::indicator:checked::after {
            content: "✓";
            color: white;
            font-size: 12px;
            font-weight: bold;
        }
        
        QTextEdit QScrollBar:vertical {
            border: none;
            background-color: #f1f3f4;
            width: 12px;
            border-radius: 6px;
            margin: 2px;
        }
        
        QTextEdit QScrollBar::handle:vertical {
            background-color: #cbd5e0;
            border-radius: 6px;
            min-height: 30px;
            margin: 2px;
        }
        
        QTextEdit QScrollBar::handle:vertical:hover {
            background-color: #a0aec0;
        }
        
        QTextEdit QScrollBar::handle:vertical:pressed {
            background-color: #718096;
        }
        
        QTextEdit QScrollBar::add-line:vertical,
        QTextEdit QScrollBar::sub-line:vertical {
            border: none;
            background: none;
            height: 0px;
        }
        
        QTextEdit QScrollBar::add-page:vertical,
        QTextEdit QScrollBar::sub-page:vertical {
            background: none;
        }
        
        QTextEdit QScrollBar:horizontal {
            border: none;
            background-color: #f1f3f4;
            height: 12px;
            border-radius: 6px;
            margin: 2px;
        }
        
        QTextEdit QScrollBar::handle:horizontal {
            background-color: #cbd5e0;
            border-radius: 6px;
            min-width: 30px;
            margin: 2px;
        }
        
        QTextEdit QScrollBar::handle:horizontal:hover {
            background-color: #a0aec0;
        }
        
        QTextEdit QScrollBar::handle:horizontal:pressed {
            background-color: #718096;
        }
        
        QTextEdit QScrollBar::add-line:horizontal,
        QTextEdit QScrollBar::sub-line:horizontal {
            border: none;
            background: none;
            width: 0px;
        }
        
        QTextEdit QScrollBar::add-page:horizontal,
        QTextEdit QScrollBar::sub-page:horizontal {
            background: none;
        }
        
        /* Custom Scrollbar Styles for QComboBox */
        QComboBox QScrollBar:vertical {
            border: none;
            background-color: #f1f3f4;
            width: 12px;
            border-radius: 6px;
            margin: 2px;
        }

        QComboBox QScrollBar::handle:vertical {
            background-color: #cbd5e0;
            border-radius: 6px;
            min-height: 30px;
            margin: 2px;
        }

        QComboBox QScrollBar::handle:vertical:hover {
            background-color: #a0aec0;
        }

        QComboBox QScrollBar::handle:vertical:pressed {
            background-color: #718096;
        }

        QComboBox QScrollBar::add-line:vertical,
        QComboBox QScrollBar::sub-line:vertical {
            border: none;
            background: none;
            height: 0px;
        }

        QComboBox QScrollBar::add-page:vertical,
        QComboBox QScrollBar::sub-page:vertical {
            background: none;
        }
        
        /* Global Scrollbar Styles (fallback) */
        QScrollBar:vertical {
            border: none;
            background-color: #f8f9fa;
            width: 14px;
            border-radius: 7px;
            margin: 2px;
        }
        
        QScrollBar::handle:vertical {
            background-color: #dee2e6;
            border-radius: 7px;
            min-height: 30px;
            margin: 2px;
        }
        
        QScrollBar::handle:vertical:hover {
            background-color: #ced4da;
        }
        
        QScrollBar::handle:vertical:pressed {
            background-color: #adb5bd;
        }
        
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical {
            border: none;
            background: none;
            height: 0px;
        }
        
        QScrollBar::add-page:vertical,
        QScrollBar::sub-page:vertical {
            background: none;
        }
        
        QScrollBar:horizontal {
            border: none;
            background-color: #f8f9fa;
            height: 14px;
            border-radius: 7px;
            margin: 2px;
        }
        
        QScrollBar::handle:horizontal {
            background-color: #dee2e6;
            border-radius: 7px;
            min-width: 30px;
            margin: 2px;
        }
        
        QScrollBar::handle:horizontal:hover {
            background-color: #ced4da;
        }
        
        QScrollBar::handle:horizontal:pressed {
            background-color: #adb5bd;
        }
        
        QScrollBar::add-line:horizontal,
        QScrollBar::sub-line:horizontal {
            border: none;
            background: none;
            width: 0px;
        }
        
        QScrollBar::add-page:horizontal,
        QScrollBar::sub-page:horizontal {
            background: none;
        }
        """)
    
    def select_file(self):
        """Handle PDF file selection"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select PDF File", "", 
            "PDF Files (*.pdf);;All files (*.*)"
        )
        
        if not file_path:
            return
        
        if not file_path.lower().endswith('.pdf'):
            self._show_status_message("Please select a valid PDF file.", "error")
            return
        
        # Verify it's a valid PDF file
        try: 
            with pdfplumber.open(file_path) as pdf:
                if len(pdf.pages) == 0:
                    self._show_status_message("The PDF file is empty.", "error")
                    return
                    
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.setText(f"File: {filename}")
            self.extract_btn.setEnabled(self.tesseract_path is not None)
            self.progress_bar.setValue(0)
            self.text_area.clear()
            self._clear_status_message()
            
        except Exception as e:
            logging.error(f"PDF opening error: {e}")
            self._show_status_message("Error opening PDF. Corrupted file?", "error")
    
    def extract_text(self):
        """Start text extraction"""
        if not self.file_path or not self.tesseract_path:
            return
        
        # Disable controls during extraction
        self.extract_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.save_docx_btn.setEnabled(False)
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # Clear text area
        self.text_area.clear()
        
        # Start extraction thread
        ocr_lang = self.lang_box.currentText()
        ocr_only_blank = self.ocr_blank_only.isChecked()
        
        self.extract_thread = ExtractThread(
            self.file_path, ocr_lang, self.tesseract_path, ocr_only_blank
        )
        
        # Connect signals
        self.extract_thread.finished.connect(self._on_extraction_finished)
        self.extract_thread.error.connect(self._on_extraction_error)
        self.extract_thread.progress.connect(self.progress_bar.setValue)
        self.extract_thread.status_update.connect(self._show_status_message)
        
        self.extract_thread.start()
        
        self._show_status_message("Extraction in progress...", "info")
    
    def _on_extraction_finished(self, text: str):
        """Handle extraction completion"""
        self.text_area.setPlainText(text)
        self.text_area.setReadOnly(False)
        
        # Re-enable controls
        self._enable_controls(True)
        
        # Enable save if there's text
        has_text = bool(text.strip())
        self.save_btn.setEnabled(has_text)
        self.save_docx_btn.setEnabled(has_text)
        
        self.progress_bar.setVisible(False)
        
        if has_text:
            word_count = len(text.split())
            char_count = len(text)
            self._show_status_message(
                f"Extraction completed! Words: {word_count}, Characters: {char_count}", 
                "success"
            )
        else:
            self._show_status_message("Extraction completed, but no text found.", "warning")
        
        # Cleanup thread
        if self.extract_thread:
            self.extract_thread.deleteLater()
            self.extract_thread = None
    
    def _on_extraction_error(self, error_msg: str):
        """Handle extraction errors"""
        self._enable_controls(True)
        self.save_btn.setEnabled(False)
        self.save_docx_btn.setEnabled(False)
        self.progress_bar.setVisible(False)
        self._show_status_message(f"Error: {error_msg}", "error")
        
        # Cleanup thread
        if self.extract_thread:
            self.extract_thread.deleteLater()
            self.extract_thread = None
    
    def _enable_controls(self, enabled: bool):
        """Enable/disable main controls"""
        self.extract_btn.setEnabled(enabled)
        self.select_btn.setEnabled(enabled)
        self.lang_box.setEnabled(enabled)
        self.ocr_blank_only.setEnabled(enabled)
    
    def save_text(self, format_type: str):
        """Save extracted text"""
        if not self.file_path:
            return
        
        # Generate default filename
        base_name = Path(self.file_path).stem
        
        if format_type == "docx":
            default_name = f"{base_name}_extracted.docx"
            file_filter = "Word Files (*.docx);;All files (*.*)"
        else:
            default_name = f"{base_name}_extracted.txt"
            file_filter = "Text Files (*.txt);;All files (*.*)"
        
        # Save dialog
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Extracted Text", default_name, file_filter
        )
        
        if not file_path:
            return
        
        # Confirm overwrite
        if os.path.exists(file_path):
            reply = QMessageBox.question(
                self, "Confirm Overwrite",
                f"The file '{os.path.basename(file_path)}' already exists.\nDo you want to overwrite it?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
        
        # Save
        try:
            text_content = self.text_area.toPlainText()
            
            if format_type == "docx":
                self._save_as_docx(file_path, text_content)
            else:
                self._save_as_txt(file_path, text_content)
            
            self._show_status_message(
                f"File saved: {os.path.basename(file_path)}", "success"
            )
            
        except Exception as e:
            logging.error(f"Save error: {e}")
            self._show_status_message(f"Save error: {str(e)}", "error")
    
    def _save_as_txt(self, file_path: str, content: str):
        """Save as text file"""
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(content)
    
    def _save_as_docx(self, file_path: str, content: str):
        """Save as Word document"""
        doc = Document()
        
        # Add paragraphs
        paragraphs = content.split('\n\n')
        for paragraph in paragraphs:
            if paragraph.strip():
                doc.add_paragraph(paragraph.strip())
        
        doc.save(file_path)
    
    def _show_status_message(self, message: str, msg_type: str = "info"):
        """Show status message"""
        color_map = {
            "success": "#28a745",
            "error": "#dc3545", 
            "warning": "#ffc107",
            "info": "#17a2b8"
        }
        
        color = color_map.get(msg_type, "#6c757d")
        self.status_label.setStyleSheet(f"color: {color}; font-weight: 600; font-size: 13px;")
        self.status_label.setText(message)
        
        # Auto-clear after 5 seconds for success messages
        if msg_type == "success":
            self.status_timer.start(5000)
    
    def _clear_status_message(self):
        """Clear status message"""
        self.status_label.clear()
        self.status_timer.stop()
    
    def closeEvent(self, event):
        """Handle application closure"""
        if self.extract_thread and self.extract_thread.isRunning():
            self.extract_thread.stop()
        event.accept()

def main():
    """Main function"""
    app = QApplication(sys.argv)
    app.setApplicationName("PDF Text Extractor")
    app.setApplicationVersion("2.0")
    
    # Check if already running (optional)
    window = PDFExtractorGUI()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__": 
    main()