import PyPDF2
import sys
from docx import Document
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QFileDialog, 
                             QLabel, QProgressBar, QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

class PDFExtractorGUI(QMainWindow): 
    def __init__(self):
        super().__init__()
        self.file_path = None  
        self.initUI()
        self.setWindowIcon(QIcon("C:/Users/aless/Downloads/pdf.ico"))
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
    """)
    
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
        
        layout.addLayout(file_layout)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.extract_btn)  
        layout.addWidget(self.text_area)
        layout.addWidget(self.save_status_label)
        layout.addWidget(self.save_btn)          
    
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select PDF File", "", "PDF Files (*.pdf)")
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(file_path.split('/')[-1])
            self.extract_btn.setEnabled(True)
            self.progress_bar.setValue(0)
            self.save_status_label.setText("")
            self.text_area.clear()
    
    def extract_text(self):
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            with open(self.file_path, "rb") as file:
                pdf_reader = PyPDF2.PdfReader(file)
                total_pages = len(pdf_reader.pages)
                extracted_text = ""
                
                for i, page in enumerate(pdf_reader.pages):
                    extracted_text += page.extract_text() + "\n\n"
                    self.progress_bar.setValue(int((i+1)/total_pages * 100))
                    QApplication.processEvents()
                self.text_area.setReadOnly(False)
                self.text_area.setPlainText(extracted_text)
                self.save_btn.setEnabled(True)
                
        except Exception as e:
            self.text_area.setReadOnly(True)
            QMessageBox.critical(self, "Error", f"Failed to extract text: {e}")
            
    
    def save_text(self):
        default_name = self.file_path.split('/')[-1].replace("pdf", "txt")
        file_path, selected_filter = QFileDialog.getSaveFileName(self, "Save Text File", default_name , "Text Files (*.txt);;Word Files (*.docx)")
        if file_path:
            try:
                if selected_filter == "Word Files (*.docx)" or file_path.endswith(".docx"):
                    doc = Document()
                    doc.add_paragraph(self.text_area.toPlainText())
                    if not file_path.endswith(".docx"):
                        file_path += ".docx"
                    doc.save(file_path)
                else:
                    if not file_path.endswith(".txt"):
                        file_path += ".txt"
                    with open(file_path, "w", encoding="utf-8") as file:
                        file.write(self.text_area.toPlainText())
                self.save_status_label.setText("Saved successfully!")
            except Exception as e:
                self.save_status_label.setStyleSheet("color: red; font-size: 13px;")
                self.save_status_label.setText(f"Error during saving process: {e}")
# ...existing code...
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFExtractorGUI()
    window.show()
    sys.exit(app.exec_())