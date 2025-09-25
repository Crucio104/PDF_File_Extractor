import sys
from docx import Document
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QFileDialog, 
                             QLabel, QProgressBar, QComboBox, QMessageBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
import pdfplumber
import os
import pytesseract
from PIL import Image, ImageFilter, ImageOps
import logging

logging.basicConfig(
    filename="pdf_extractor.log", 
    level=logging.ERROR,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
tesseract_path = r"C:/Program Files/Tesseract-OCR/tesseract.exe"

class ExtractThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, file_path, ocr_lang, tesseract_path): 
        super().__init__()
        self.file_path = file_path
        self.ocr_lang = ocr_lang
        self.tesseract_path = tesseract_path
    def run(self):
        try:
            with pdfplumber.open(self.file_path) as pdf:
                extracted_text = ""
                total_pages = len(pdf.pages)
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text and text.strip():
                        extracted_text += text + "\n\n"
                    else:
                        pil_image = page.to_image(resolution=300).original
                        pil_image = pil_image.convert("L")
                        pil_image = ImageOps.autocontrast(pil_image)
                        pil_image = pil_image.point(lambda x: 0 if x < 150 else 255, '1')
                        pil_image = pil_image.filter(ImageFilter.MedianFilter(size=3))
                        try:
                            pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
                            ocr_text = pytesseract.image_to_string(pil_image, lang=self.ocr_lang)
                        except pytesseract.TesseractNotFoundError:
                            self.error.emit("Tesseract not found. Please install Tesseract OCR.")
                            logging.error("Tesseract not found.")
                            return
                        except Exception as e:
                            if "Failed loading language" in str(e):
                                self.error.emit(f"Language file '{self.ocr_lang}' not found in tessdata.")
                                logging.error(f"Language file '{self.ocr_lang}' not found in tessdata.")
                                return
                            logging.error(str(e))
                            raise
                        extracted_text += ocr_text + "\n\n"
                    self.progress.emit(int((i + 1) / total_pages * 100))
            self.finished.emit(extracted_text)
        except Exception as e:
            print(str(e))
            self.error.emit(str(e))


class PDFExtractorGUI(QMainWindow): 
    def __init__(self):
        super().__init__()
        self.file_path = None  
        self.initUI()
        icon_path = "C:/Users/aless/Downloads/pdf.ico"
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            self.setWindowIcon(QIcon())
        self.setStyleSheet("""
        QMainWindow {
            background-color: #f5f5f7;
        }
        QWidget {
            background-color: #f5f5f7;
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        QLabel {
            color: #2c3e50;
            font-size: 15px;
            padding: 5px;
        }
        
        QTextEdit QScrollBar:vertical {
        border: none;
        background-color: #f1f3f4;
        width: 14px;
        border-radius: 7px;
        margin: 2px;
        }
        
        QTextEdit QScrollBar::handle:vertical {
            background-color: #cbd5e0;
            border-radius: 7px;
            min-height: 40px;
        }
        
        QTextEdit QScrollBar::handle:vertical:hover {
            background-color: #a0aec0;
        }
        
        QTextEdit QScrollBar::handle:vertical:pressed {
            background-color: #718096;
        }
        
        QTextEdit QScrollBar::add-line:vertical, 
        QTextEdit QScrollBar::sub-line:vertical {
            height: 0px;
        }
        
        QTextEdit QScrollBar::add-page:vertical, 
        QTextEdit QScrollBar::sub-page:vertical {
            background: none;
        }
        
        QPushButton {
            background-color: #3498db;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 6px;
            font-weight: 500;
            font-size: 15px;
        }
        QPushButton:hover {
            background-color: #2980b9;
        }
        QPushButton:pressed {
            background-color: #21618c;
        }
        QPushButton:disabled {
            background-color: #bdc3c7;
            color: #7f8c8d;
        }
        QTextEdit {
            background-color: white;
            border: 2px solid #dce4ec;
            border-radius: 6px;
            padding: 8px;
            font-size: 13px;
            color: #2c3e50;
        }
        QTextEdit:disabled {
            background-color: #ecf0f1;
            color: #7f8c8d;
        }
        QProgressBar {
            border: 2px solid #dce4ec;
            border-radius: 4px;
            text-align: center;
            background-color: white;
            color: #2c3e50;
        }
        QProgressBar::chunk {
            background-color: #228b22;
            border-radius: 3px;
        }
        QComboBox {
        background-color: white;
        border: 2px solid #dce4ec;
        border-radius: 6px;
        padding: 6px 12px;
        font-size: 15px;
        color: #2c3e50;
        font-family: 'Segoe UI', Arial, sans-serif;
        }
        QComboBox:focus {
            border: 2px solid #3498db;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 28px;
            border-left: 2px solid #dce4ec;
            border-top-right-radius: 6px;
            border-bottom-right-radius: 6px;
            background: #f1f3f4;
        }
        QComboBox::down-arrow {
            image: none;
            width: 0;
            height: 0;
        }
        QComboBox QAbstractItemView {
            background: white;
            border: 2px solid #dce4ec;
            selection-background-color: #3498db;
            selection-color: white;
            font-size: 15px;
        }
        QComboBox QScrollBar:vertical {
        border: none;
        background-color: #f1f3f4;
        width: 14px;
        border-radius: 7px;
        margin: 2px;
        }

        QComboBox QScrollBar::handle:vertical {
            background-color: #cbd5e0;
            border-radius: 7px;
            min-height: 40px;
        }

        QComboBox QScrollBar::handle:vertical:hover {
            background-color: #a0aec0;
        }

        QComboBox QScrollBar::handle:vertical:pressed {
            background-color: #718096;
        }

        QComboBox QScrollBar::add-line:vertical,
        QComboBox QScrollBar::sub-line:vertical {
            height: 0px;
        }

        QComboBox QScrollBar::add-page:vertical,
        QComboBox QScrollBar::sub-page:vertical {
            background: none;
        }

    """)
    
    def get_available_languages(self):
        tessdata_dir = r"C:/Program Files/Tesseract-OCR/tessdata"
        langs = []
        if os.path.exists(tessdata_dir):
            for f in os.listdir(tessdata_dir):
                if f.endswith(".traineddata"):
                    langs.append(f.split(".")[0])
            return sorted(langs) if langs else ["eng"]
        else:
            logging.error("Tesseract-OCR installation not found.")
            return ["eng"]

    def initUI(self):
        self.setWindowTitle("PDF Text Extractor")
        self.setGeometry(100, 100, 600, 600)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        self.select_btn = QPushButton("Select PDF")
        self.select_btn.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.select_btn)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setMaximumHeight(12)
        self.progress_bar.setFormat("")
        self.progress_bar.setVisible(True)
        
        

        self.lang_box = QComboBox()
        self.lang_box.addItems(self.get_available_languages())
        self.lang_box.setCurrentText("eng")
        lang_layout = QHBoxLayout()
        lang_label = QLabel("Select OCR Language:")
        lang_layout.addWidget(lang_label)
        lang_layout.addWidget(self.lang_box)
        
        self.extract_btn = QPushButton("Extract Text")
        self.extract_btn.clicked.connect(self.extract_text)
        self.extract_btn.setEnabled(False) 
        
        self.text_area = QTextEdit()
        self.text_area.setPlaceholderText("Extracted text will appear here...")
        self.text_area.setReadOnly(True)
        
        self.save_status_label = QLabel("")
        self.save_status_label.setStyleSheet("color: green; font-size: 13px;")
        self.save_status_label.setAlignment(Qt.AlignCenter)
        
        self.save_btn = QPushButton("Save Text")
        self.save_btn.clicked.connect(self.save_text)
        self.save_btn.setEnabled(False)  
        
        self.select_btn.setToolTip("Select a PDF file to extract text")
        self.extract_btn.setToolTip("Extract text from the selected PDF")
        self.save_btn.setToolTip("Save the extracted text")
        self.lang_box.setToolTip("Select the OCR language for scanned pages")
                
        layout.addLayout(file_layout)
        layout.addLayout(lang_layout)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.extract_btn)  
        layout.addWidget(self.text_area)
        layout.addWidget(self.save_status_label)
        layout.addWidget(self.save_btn)          
    
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select PDF File", "", "PDF Files (*.pdf)")
        
        if file_path and file_path.lower().endswith('.pdf'):
            try: 
                with pdfplumber.open(file_path) as pdf:
                    pass
                self.file_path = file_path
                self.file_label.setText(os.path.basename(file_path))
                self.extract_btn.setEnabled(True)
                self.progress_bar.setValue(0)
                self.save_status_label.setText("")
                self.text_area.clear()
            except Exception:
                self.save_status_label.setStyleSheet("color: red; font-size: 13px;")
                self.save_status_label.setText("Select a valid PDF file.")
        else:
            self.save_status_label.setStyleSheet("color: red; font-size: 13px;")
            self.save_status_label.setText("Select a valid PDF file.")
    
    def extract_text(self):
        self.extract_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.text_area.clear()
        self.save_btn.setEnabled(False)
        self.save_status_label.setText("Extracting...")
        ocr_lang = self.lang_box.currentText()
        self.thread = ExtractThread(self.file_path, ocr_lang, tesseract_path)
        self.thread.finished.connect(self.on_extraction_finished)
        self.thread.error.connect(self.on_extraction_error)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.start()
        
    def on_extraction_finished(self, text):
        self.text_area.setPlainText(text)
        self.text_area.setReadOnly(False)
        self.save_btn.setEnabled(bool(text.strip()))
        self.save_status_label.setStyleSheet("color: green; font-size: 13px;")
        self.save_status_label.setText("Extraction completed!")
        self.extract_btn.setEnabled(True)
        self.select_btn.setEnabled(True)

    def on_extraction_error(self, error_msg):
        self.text_area.setReadOnly(True)
        self.save_btn.setEnabled(False)
        self.save_status_label.setStyleSheet("color: red; font-size: 13px;")
        self.save_status_label.setText(f"Error: {error_msg}")
        self.extract_btn.setEnabled(True)
        self.select_btn.setEnabled(True)

    def save_text(self):
        base_name = os.path.splitext(os.path.basename(self.file_path))[0]
        default_name = base_name + ".txt"
        file_path, selected_filter = QFileDialog.getSaveFileName(self, "Save Text File", default_name , "Text Files (*.txt);;Word Files (*.docx)")
        if file_path:
            if os.path.exists(file_path):
                reply = QMessageBox.question(
                    self,
                    "Overwrite?",
                    f"{file_path} exists. Overwrite?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.No:
                    self.save_status_label.setStyleSheet("color: orange; font-size: 13px;")
                    self.save_status_label.setText("Save operation canceled.")
                    return
            try:
                if selected_filter == "Word Files (*.docx)" or file_path.endswith(".docx"):
                    doc = Document()
                    for paragraph in self.text_area.toPlainText().split('\n\n'):
                        doc.add_paragraph(paragraph)
                    if not file_path.endswith(".docx"):
                        file_path += ".docx"
                    doc.save(file_path)
                else:
                    if not file_path.endswith(".txt"):
                        file_path += ".txt"
                    with open(file_path, "w", encoding="utf-8") as file:
                        file.write(self.text_area.toPlainText())
                self.save_status_label.setStyleSheet("color: green; font-size: 13px;")
                self.save_status_label.setText("Saved successfully!")
            except Exception as e:
                self.save_status_label.setStyleSheet("color: red; font-size: 13px;")
                self.save_status_label.setText(f"Error during saving process: {e}")


if __name__ == "__main__": 
    app = QApplication(sys.argv)
    window = PDFExtractorGUI()
    window.show()
    sys.exit(app.exec_())