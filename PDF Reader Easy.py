import PyPDF2
import re
import os
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QFileDialog, 
                             QLabel, QProgressBar, QMessageBox)
from PyQt5.QtCore import Qt

class PDFExtractorGUI(QMainWindow):  # CORRETTO: rimuovi (self)
    def __init__(self):
        super().__init__()
        self.file_path = None  # Aggiungi questa linea
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle("PDF Text Extractor")
        self.setGeometry(100, 100, 800, 600)
        
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
        self.progress_bar.setVisible(False)
        
        # AGGIUNGI QUESTO PULSANTE MANCANTE:
        self.extract_btn = QPushButton("Extract Text")
        self.extract_btn.clicked.connect(self.extract_text)
        self.extract_btn.setEnabled(False)  # Disabilitato inizialmente
        
        self.text_area = QTextEdit()
        self.text_area.setPlaceholderText("Extracted text will appear here...")
        
        self.save_btn = QPushButton("Save Text")
        self.save_btn.clicked.connect(self.save_text)
        self.save_btn.setEnabled(False)  # Disabilitato inizialmente
        
        layout.addLayout(file_layout)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.extract_btn)  # AGGIUNGI QUESTO
        layout.addWidget(self.text_area)
        layout.addWidget(self.save_btn)          
    
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select PDF File", "", "PDF Files (*.pdf)")
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(file_path.split('/')[-1])
            self.extract_btn.setEnabled(True)
    
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
                
                self.text_area.setPlainText(extracted_text)
                self.save_btn.setEnabled(True)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract text: {e}")
            
        finally:
            self.progress_bar.setVisible(False)
    
    def save_text(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Text File", "", "Text Files (*.txt)")
        
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as file:
                    file.write(self.text_area.toPlainText())
                QMessageBox.information(self, "Success", "Text saved successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFExtractorGUI()
    window.show()
    sys.exit(app.exec_())