import sys
import os
import logging
import ctypes
import shutil
from pathlib import Path
from typing import Optional, List
from docx import Document
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QFileDialog, 
                             QLabel, QProgressBar, QComboBox, QMessageBox, QCheckBox, QAction, QGroupBox, QFrame)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon, QFont
import pdfplumber
import pytesseract
from PIL import Image, ImageFilter, ImageOps, ImageDraw, ImageFont

logging.basicConfig(
    filename="pdf_extractor.log", 
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    filemode='a'
)

class TesseractConfig:
    
    
    @staticmethod
    def find_tesseract_path() -> Optional[str]:
        
        possible_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            "/usr/bin/tesseract",
            "/usr/local/bin/tesseract",
            "/opt/homebrew/bin/tesseract"
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        import shutil
        tesseract_cmd = shutil.which("tesseract")
        if tesseract_cmd:
            return tesseract_cmd
            
        return None
    
    @staticmethod
    def get_available_languages(tesseract_path: str) -> List[str]:
    
        try:
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            try:
                langs = pytesseract.get_languages(config="")
            except TypeError:
                langs = pytesseract.get_languages()
            return sorted(langs) if langs else ["eng"]
        except Exception as e:
            logging.error(f"Error retrieving languages: {e}")
            return ["eng"]

class ExtractThread(QThread):
    
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)
    status_update = pyqtSignal(str)

    def __init__(self, file_path: str, ocr_lang: str, tesseract_path: str, 
                 ocr_only_blank: bool = True, fast_mode: bool = True): 
        super().__init__()
        self.file_path = file_path
        self.ocr_lang = ocr_lang
        self.tesseract_path = tesseract_path
        self.ocr_only_blank = ocr_only_blank
        self.fast_mode = fast_mode
        self._is_running = True
    
    def stop(self):
        
        self._is_running = False
        self.quit()
        self.wait()
    
    def run(self):
        try:
            ocr_enabled = bool(self.tesseract_path)
            if ocr_enabled:
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
            
            with pdfplumber.open(self.file_path) as pdf:
                extracted_text = ""
                total_pages = len(pdf.pages)
                
                self.status_update.emit(f"Processing {total_pages} pages...")
                
                for i, page in enumerate(pdf.pages):
                    if not self._is_running:
                        return
                    
                    text = page.extract_text()
                    has_text = text and text.strip()
                    
                    if has_text and self.ocr_only_blank:
                        extracted_text += text + "\n\n"
                        self.status_update.emit(f"Page {i+1}: text extracted directly")
                    else:
                        if ocr_enabled:
                            self.status_update.emit(f"Page {i+1}: performing OCR...")
                            try:
                                ocr_text = self._perform_ocr(page)
                                extracted_text += ocr_text + "\n\n"
                            except Exception as e:
                                logging.error(f"OCR error on page {i+1}: {e}")
                                if has_text:
                                    extracted_text += text + "\n\n"
                                else:
                                    extracted_text += f"[Error processing page {i+1}]\n\n"
                        else:
                            if has_text:
                                extracted_text += text + "\n\n"
                                self.status_update.emit(f"Page {i+1}: OCR unavailable, text extracted directly")
                            else:
                                extracted_text += f"[OCR unavailable; no text found on page {i+1}]\n\n"
                                self.status_update.emit(f"Page {i+1}: OCR unavailable; skipping empty text")
                    
                    progress = int((i + 1) / total_pages * 100)
                    self.progress.emit(progress)
                
            self.finished.emit(extracted_text)
            
        except FileNotFoundError:
            self.error.emit("PDF file not found.")
        except Exception as e:
            logging.error(f"Error during extraction: {e}")
            self.error.emit(f"Error during extraction: {str(e)}")
    
    def _perform_ocr(self, page) -> str:
        
        try:
            resolution = 200 if self.fast_mode else 300
            pil_image = page.to_image(resolution=resolution).original
            pil_image = self._preprocess_image(pil_image)
            ocr_text = pytesseract.image_to_string(pil_image, lang=self.ocr_lang)
            pil_image.close()
            
            return ocr_text
            
        except pytesseract.TesseractNotFoundError:
            raise Exception("Tesseract not found. Please check installation.")
        except Exception as e:
            if "Failed loading language" in str(e):
                raise Exception(f"Language file '{self.ocr_lang}' not found.")
            raise e
    
    def _preprocess_image(self, image: Image.Image) -> Image.Image:
        
        image = image.convert("L")
        image = ImageOps.autocontrast(image)
        if not self.fast_mode:
            image = image.filter(ImageFilter.MedianFilter(size=3))
        image = image.point(lambda x: 0 if x < 150 else 255)

        return image

class PDFExtractorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_path: Optional[str] = None  
        self.extract_thread: Optional[ExtractThread] = None
        self.tesseract_path: Optional[str] = None
        self.dark_mode: bool = True
        self.fast_mode: bool = True
        
        self._init_tesseract()
        self.initUI()
        self._setup_styles()
        self._set_icon()
        self._apply_dark_titlebar()

        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self._clear_status_message)
    
    def _init_tesseract(self):
        self.tesseract_path = TesseractConfig.find_tesseract_path()
        if not self.tesseract_path:
            QMessageBox.warning(
                self,
                "Tesseract not found", 
                "Tesseract OCR was not found. Please install it to use OCR functionality."
            )
    
    def _set_icon(self):
        base_dir = os.path.dirname(__file__)
        assets_dir = os.path.join(base_dir, "assets")
        icon_path = os.path.join(assets_dir, "app_icon.png")
        try:
            self._ensure_icon_asset(icon_path)
            self.setWindowIcon(QIcon(icon_path))
        except Exception:
            self.setWindowIcon(self.style().standardIcon(self.style().SP_FileDialogDetailedView))

    def _ensure_icon_asset(self, icon_path: str):
        assets_dir = os.path.dirname(icon_path)
        if not os.path.exists(assets_dir):
            os.makedirs(assets_dir, exist_ok=True)
        if os.path.exists(icon_path):
            return
        size = (256, 256)
        img = Image.new("RGBA", size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        margin = 36
        page_rect = [margin, margin, size[0] - margin, size[1] - margin]
        draw.rounded_rectangle(page_rect, radius=24, fill=(255, 255, 255, 255), outline=(220, 220, 220, 255), width=6)
        fold_size = 56
        fx, fy = size[0] - margin - fold_size, margin
        tri = [(size[0] - margin, fy), (size[0] - margin, fy + fold_size), (fx, fy)]
        draw.polygon(tri, fill=(240, 240, 240, 255))
        bar_height = 44
        bar_rect = [margin + 24, size[1] - margin - bar_height - 24, size[0] - margin - 24, size[1] - margin - 24]
        draw.rounded_rectangle(bar_rect, radius=18, fill=(220, 53, 69, 255))
        try:
            font = ImageFont.truetype("arial.ttf", 64)
        except Exception:
            font = ImageFont.load_default()
        text = "PDF"
        tw, th = draw.textsize(text, font=font)
        tx = (size[0] - tw) // 2
        ty = (size[1] // 2) - (th // 2)
        draw.text((tx, ty), text, fill=(220, 53, 69, 255), font=font)
        img.save(icon_path, format="PNG")

    def _apply_dark_titlebar(self):
        if sys.platform == "win32":
            try:
                hwnd = int(self.winId())
                DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                value = ctypes.c_int(1)
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    DWMWA_USE_IMMERSIVE_DARK_MODE,
                    ctypes.byref(value),
                    ctypes.sizeof(value)
                )
            except Exception:
                pass

    def initUI(self):
        self.setWindowTitle("PDF Text Extractor v2.0")
        self.setGeometry(100, 100, 650, 680)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        menubar = self.menuBar()
        file_menu = menubar.addMenu("File")
        help_menu = menubar.addMenu("Help")

        self.actionOpen = QAction("Open PDF", self)
        self.actionOpen.setShortcut("Ctrl+O")
        self.actionOpen.triggered.connect(self.select_file)
        file_menu.addAction(self.actionOpen)

        self.actionSaveTxt = QAction("Save as TXT", self)
        self.actionSaveTxt.setShortcut("Ctrl+S")
        self.actionSaveTxt.triggered.connect(lambda: self.save_text("txt"))
        self.actionSaveTxt.setEnabled(False)
        file_menu.addAction(self.actionSaveTxt)

        self.actionSaveDocx = QAction("Save as DOCX", self)
        self.actionSaveDocx.triggered.connect(lambda: self.save_text("docx"))
        self.actionSaveDocx.setEnabled(False)
        file_menu.addAction(self.actionSaveDocx)

        self.actionExit = QAction("Exit", self)
        self.actionExit.setShortcut("Ctrl+Q")
        self.actionExit.triggered.connect(self.close)
        file_menu.addAction(self.actionExit)

        actionAbout = QAction("About", self)
        actionAbout.triggered.connect(lambda: QMessageBox.information(self, "About", "PDF Text Extractor\nOCR with Tesseract, text via pdfplumber."))
        help_menu.addAction(actionAbout)
        
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        self._create_file_selection_section(layout)
        self._create_ocr_config_section(layout)
        self._create_progress_section(layout)
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)
        self.extract_btn = QPushButton("🚀 Extract Text")
        self.extract_btn.clicked.connect(self.extract_text)
        self.extract_btn.setEnabled(False)
        self.extract_btn.setMinimumHeight(38)
        buttons_layout.addWidget(self.extract_btn)
        self.cancel_btn = QPushButton("❌ Cancel")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel_extraction)
        self.cancel_btn.setMinimumHeight(38)
        buttons_layout.addWidget(self.cancel_btn)
        layout.addLayout(buttons_layout)
        
        self._create_text_area_section(layout)
        self._create_save_section(layout)
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)
    
    def _create_file_selection_section(self, layout: QVBoxLayout):
        group = QGroupBox("📄 PDF File")
        group.setStyleSheet("QGroupBox { font-weight: bold; font-size: 14px; padding-top: 8px; margin-top: 8px; }")
        file_layout = QHBoxLayout()
        
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("font-weight: normal; padding: 6px;")
        self.select_btn = QPushButton("📁 Select PDF")
        self.select_btn.clicked.connect(self.select_file)
        
        file_layout.addWidget(self.file_label, 1)
        file_layout.addWidget(self.select_btn)
        group.setLayout(file_layout)
        layout.addWidget(group)
    
    def _create_ocr_config_section(self, layout: QVBoxLayout):
        group = QGroupBox("⚙️ OCR Settings")
        group.setStyleSheet("QGroupBox { font-weight: bold; font-size: 14px; padding-top: 8px; margin-top: 8px; }")
        ocr_layout = QVBoxLayout()
        ocr_layout.setSpacing(10)
        
        lang_layout = QHBoxLayout()
        lang_label = QLabel("Language:")
        lang_label.setMinimumWidth(80)
        self.lang_box = QComboBox()
        self.lang_box.setMaxVisibleItems(10)
        self.lang_box.view().setVerticalScrollMode(self.lang_box.view().ScrollPerPixel)
        
        if self.tesseract_path:
            languages = TesseractConfig.get_available_languages(self.tesseract_path)
            self.lang_box.addItems(languages)
            if "eng" in languages:
                self.lang_box.setCurrentText("eng")
        else:
            self.lang_box.addItem("OCR not available")
            self.lang_box.setEnabled(False)
        
        lang_layout.addWidget(lang_label)
        lang_layout.addWidget(self.lang_box, 1)
        
        self.ocr_blank_only = QCheckBox("🔍 Use OCR only on pages without text")
        self.ocr_blank_only.setChecked(True)
        self.ocr_blank_only.setToolTip("If unchecked, will use OCR on all pages")

        self.fast_mode_box = QCheckBox("⚡ Fast Mode (lower DPI, minimal preprocessing)")
        self.fast_mode_box.setChecked(True)
        self.fast_mode_box.stateChanged.connect(lambda _: setattr(self, 'fast_mode', self.fast_mode_box.isChecked()))
        
        ocr_layout.addLayout(lang_layout)
        ocr_layout.addWidget(self.ocr_blank_only)
        ocr_layout.addWidget(self.fast_mode_box)
        group.setLayout(ocr_layout)
        layout.addWidget(group)
    
    def _create_progress_section(self, layout: QVBoxLayout):
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setMaximumHeight(20)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
    
    def _create_text_area_section(self, layout: QVBoxLayout):
        group = QGroupBox("📝 Extracted Text")
        group.setStyleSheet("QGroupBox { font-weight: bold; font-size: 14px; padding-top: 8px; margin-top: 8px; }")
        text_layout = QVBoxLayout()
        
        self.text_area = QTextEdit()
        self.text_area.setPlaceholderText("📄 Extracted text will appear here...")
        self.text_area.setReadOnly(True)
        
        font = QFont("Consolas", 10)
        if not font.exactMatch():
            font = QFont("Courier New", 10)
        self.text_area.setFont(font)
        
        text_layout.addWidget(self.text_area)
        group.setLayout(text_layout)
        layout.addWidget(group)
    
    def _create_save_section(self, layout: QVBoxLayout):
        save_layout = QHBoxLayout()
        save_layout.setSpacing(8)
        
        self.save_btn = QPushButton("💾 Save as TXT")
        self.save_btn.clicked.connect(lambda: self.save_text("txt"))
        self.save_btn.setEnabled(False)
        self.save_btn.setMinimumHeight(36)
        
        self.save_docx_btn = QPushButton("📄 Save as DOCX")
        self.save_docx_btn.clicked.connect(lambda: self.save_text("docx"))
        self.save_docx_btn.setEnabled(False)
        self.save_docx_btn.setMinimumHeight(36)
        
        save_layout.addWidget(self.save_btn)
        save_layout.addWidget(self.save_docx_btn)
        layout.addLayout(save_layout)
    
    def _setup_styles(self):
        self.apply_theme(self.dark_mode)
    
    def apply_theme(self, dark: bool):
        if dark:
            self.setStyleSheet("""
        QMainWindow { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1a1a2e, stop:1 #16213e); }
        QWidget { background-color: transparent; color: #e6e6e6; font-family: 'Segoe UI', 'Arial', sans-serif; font-size: 14px; }
        QMenuBar { 
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2a2a3e, stop:1 #1f1f32);
            color: #ffffff;
            border-bottom: 2px solid #4cc9f0;
            padding: 2px 8px;
            font-size: 14px;
            font-weight: 600;
        }
        QMenuBar::item { 
            background-color: transparent;
            padding: 4px 14px;
            margin: 1px 3px;
            border-radius: 5px;
        }
        QMenuBar::item:selected { 
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4361ee, stop:1 #3a0ca3);
            color: #ffffff;
        }
        QMenuBar::item:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3a0ca3, stop:1 #240046);
        }
        QMenu { 
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2a2a3e, stop:1 #252526);
            color: #e6e6e6;
            border: 2px solid #4cc9f0;
            border-radius: 8px;
            padding: 8px;
        }
        QMenu::item {
            padding: 10px 24px;
            margin: 2px 4px;
            border-radius: 6px;
        }
        QMenu::item:selected { 
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4361ee, stop:1 #3a0ca3);
            color: #ffffff;
        }
        QMenu::separator {
            height: 2px;
            background-color: #4cc9f0;
            margin: 6px 12px;
        }
        QGroupBox { color: #4cc9f0; border: 2px solid #3c3c3c; border-radius: 7px; margin-top: 6px; padding: 12px; background-color: rgba(37, 37, 38, 0.6); font-weight: bold; font-size: 14px; }
        QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; }
        QLabel { color: #e6e6e6; font-size: 13px; padding: 4px; }
        QPushButton { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4361ee, stop:1 #3a0ca3); color: white; border: none; padding: 10px 20px; border-radius: 7px; font-weight: 600; font-size: 14px; }
        QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4cc9f0, stop:1 #4361ee); }
        QPushButton:pressed { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3a0ca3, stop:1 #240046); }
        QPushButton:disabled { background-color: #5a5a5a; color: #bfbfbf; }
        QTextEdit { background-color: #252526; border: 2px solid #4cc9f0; border-radius: 7px; padding: 10px; font-size: 12px; color: #e6e6e6; selection-background-color: #4361ee; }
        QProgressBar { border: 2px solid #3c3c3c; border-radius: 7px; text-align: center; background-color: #252526; color: #4cc9f0; font-weight: 600; height: 22px; }
        QProgressBar::chunk { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4361ee, stop:1 #4cc9f0); border-radius: 5px; }
        QCheckBox { color: #e6e6e6; font-size: 13px; spacing: 8px; }
        QCheckBox::indicator { width: 18px; height: 18px; border: 2px solid #4cc9f0; border-radius: 4px; background-color: #252526; }
        QCheckBox::indicator:checked { background-color: #4361ee; border-color: #4cc9f0; }
        QComboBox {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2e2e42, stop:1 #252526);
            border: 2px solid #4cc9f0;
            border-radius: 6px;
            padding: 5px 10px;
            color: #ffffff;
            font-size: 13px;
            font-weight: 500;
            min-height: 22px;
        }
        QComboBox:hover {
            border: 2px solid #4cc9f0;
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3a3a50, stop:1 #2e2e42);
            color: #4cc9f0;
        }
        QComboBox:focus {
            border: 2px solid #4361ee;
        }
        QComboBox:disabled {
            background-color: #3a3a3a;
            border: 2px solid #5a5a5a;
            color: #8a8a8a;
        }
        QComboBox::drop-down {
            border: none;
            width: 26px;
        }
        QComboBox::down-arrow {
            image: none;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-top: 5px solid #4cc9f0;
            margin-right: 6px;
        }
        QComboBox::down-arrow:hover {
            border-top-color: #ffffff;
        }
        QComboBox QAbstractItemView {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2e2e42, stop:1 #1f1f32);
            border: 2px solid #4cc9f0;
            border-radius: 6px;
            selection-background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4cc9f0, stop:1 #4361ee);
            selection-color: #ffffff;
            color: #e6e6e6;
            padding: 4px;
            outline: none;
        }
        QComboBox QAbstractItemView::item {
            padding: 8px 12px;
            border-radius: 5px;
            margin: 1px 3px;
            min-height: 20px;
        }
        QComboBox QAbstractItemView::item:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4cc9f0, stop:1 #4361ee);
            color: #ffffff;
        }
        QScrollBar:vertical { background-color: #1e1e1e; width: 12px; border-radius: 6px; }
        QScrollBar::handle:vertical { background-color: #4cc9f0; border-radius: 6px; min-height: 30px; }
        QScrollBar::handle:vertical:hover { background-color: #4361ee; }
            """)
        else:
            self.setStyleSheet("""
        QMainWindow {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #e0f7fa, stop:1 #f1f8e9);
        }
        QWidget {
            background-color: transparent;
            font-family: 'Segoe UI', 'Arial', sans-serif;
            font-size: 14px;
        }
        QMenuBar {
            background-color: #ffffff;
            color: #2c3e50;
            border-bottom: 1px solid #bbdefb;
        }
        QMenuBar::item {
            background-color: transparent;
            padding: 4px 12px;
        }
        QMenuBar::item:selected {
            background-color: #42a5f5;
            color: white;
        }
        QMenu {
            background-color: #ffffff;
            color: #2c3e50;
            border: 1px solid #bbdefb;
        }
        QMenu::item:selected {
            background-color: #42a5f5;
            color: white;
        }
        QGroupBox {
            color: #1976d2;
            border: 2px solid #bbdefb;
            border-radius: 8px;
            margin-top: 8px;
            padding: 15px;
            background-color: rgba(255, 255, 255, 0.8);
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 15px;
            padding: 0 8px;
        }
        QLabel {
            color: #2c3e50;
            font-size: 14px;
            padding: 5px;
        }
        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #42a5f5, stop:1 #1976d2);
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 8px;
            font-weight: 600;
            font-size: 15px;
        }
        QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #66bb6a, stop:1 #43a047);
        }
        QPushButton:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #1976d2, stop:1 #0d47a1);
        }
        QPushButton:disabled {
            background-color: #bdbdbd;
            color: #757575;
        }
        QTextEdit {
            background-color: white;
            border: 2px solid #42a5f5;
            border-radius: 8px;
            padding: 12px;
            font-size: 13px;
            color: #2c3e50;
            selection-background-color: #bbdefb;
        }
        QProgressBar {
            border: 2px solid #bbdefb;
            border-radius: 8px;
            text-align: center;
            background-color: white;
            color: #1976d2;
            font-weight: 600;
            height: 24px;
        }
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #42a5f5, stop:1 #66bb6a);
            border-radius: 6px;
        }
        QCheckBox {
            color: #2c3e50;
            font-size: 14px;
            spacing: 10px;
        }
        QCheckBox::indicator {
            width: 20px;
            height: 20px;
            border: 2px solid #42a5f5;
            border-radius: 4px;
            background-color: white;
        }
        QCheckBox::indicator:checked {
            background-color: #42a5f5;
            border-color: #1976d2;
        }
        QCheckBox::indicator:hover {
            border-color: #66bb6a;
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

    def toggle_dark_mode(self, checked: bool):
        self.dark_mode = checked
        self.apply_theme(self.dark_mode)
    
    def select_file(self):
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select PDF File", "", 
            "PDF Files (*.pdf);;All files (*.*)"
        )
        
        if not file_path:
            return
        
        if not file_path.lower().endswith('.pdf'):
            self._show_status_message("Please select a valid PDF file.", "error")
            return
        
        try: 
            with pdfplumber.open(file_path) as pdf:
                if len(pdf.pages) == 0:
                    self._show_status_message("The PDF file is empty.", "error")
                    return
                    
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.setText(f"File: {filename}")
            self.extract_btn.setEnabled(True)
            self.progress_bar.setValue(0)
            self.text_area.clear()
            self._clear_status_message()
            self.actionSaveTxt.setEnabled(False)
            self.actionSaveDocx.setEnabled(False)
            
        except Exception as e:
            logging.error(f"PDF opening error: {e}")
            self._show_status_message("Error opening PDF. Corrupted file?", "error")
    
    def extract_text(self):
        
        if not self.file_path:
            return
        
        self.extract_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.save_docx_btn.setEnabled(False)
        self.actionSaveTxt.setEnabled(False)
        self.actionSaveDocx.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        self.text_area.clear()
        
        ocr_lang = self.lang_box.currentText()
        ocr_only_blank = self.ocr_blank_only.isChecked()
        
        self.extract_thread = ExtractThread(
            self.file_path, ocr_lang, self.tesseract_path, ocr_only_blank, self.fast_mode
        )
        
        self.extract_thread.finished.connect(self._on_extraction_finished)
        self.extract_thread.error.connect(self._on_extraction_error)
        self.extract_thread.progress.connect(self.progress_bar.setValue)
        self.extract_thread.status_update.connect(self._show_status_message)
        
        self.extract_thread.start()
        
        self._show_status_message("Extraction in progress...", "info")
    
    def _on_extraction_finished(self, text: str):
        
        self.text_area.setPlainText(text)
        self.text_area.setReadOnly(False)
        
        self._enable_controls(True)
        self.cancel_btn.setEnabled(False)
        
        has_text = bool(text.strip())
        self.save_btn.setEnabled(has_text)
        self.save_docx_btn.setEnabled(has_text)
        self.actionSaveTxt.setEnabled(has_text)
        self.actionSaveDocx.setEnabled(has_text)
        
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
        
        if self.extract_thread:
            self.extract_thread.deleteLater()
            self.extract_thread = None
    
    def _on_extraction_error(self, error_msg: str):
        
        self._enable_controls(True)
        self.save_btn.setEnabled(False)
        self.save_docx_btn.setEnabled(False)
        self.actionSaveTxt.setEnabled(False)
        self.actionSaveDocx.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        self.progress_bar.setVisible(False)
        self._show_status_message(f"Error: {error_msg}", "error")
        
        if self.extract_thread:
            self.extract_thread.deleteLater()
            self.extract_thread = None
    
    def _enable_controls(self, enabled: bool):
        
        can_extract = enabled and (self.file_path is not None)
        self.extract_btn.setEnabled(can_extract)
        self.select_btn.setEnabled(enabled)
        self.lang_box.setEnabled(self.tesseract_path is not None)
        self.ocr_blank_only.setEnabled(enabled)

    def cancel_extraction(self):
        if self.extract_thread and self.extract_thread.isRunning():
            self.extract_thread.stop()
            self._show_status_message("Extraction cancelled.", "warning")
        self._enable_controls(True)
        self.cancel_btn.setEnabled(False)
        self.progress_bar.setVisible(False)
    
    def save_text(self, format_type: str):
        
        if not self.file_path:
            return
        
        base_name = Path(self.file_path).stem
        
        if format_type == "docx":
            default_name = f"{base_name}_extracted.docx"
            file_filter = "Word Files (*.docx);;All files (*.*)"
        else:
            default_name = f"{base_name}_extracted.txt"
            file_filter = "Text Files (*.txt);;All files (*.*)"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Extracted Text", default_name, file_filter
        )
        
        if not file_path:
            return
        
        if os.path.exists(file_path):
            reply = QMessageBox.question(
                self, "Confirm Overwrite",
                f"The file '{os.path.basename(file_path)}' already exists.\nDo you want to overwrite it?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
        
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
        
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(content)
    
    def _save_as_docx(self, file_path: str, content: str):
        
        doc = Document()
        
        paragraphs = content.split('\n\n')
        for paragraph in paragraphs:
            if paragraph.strip():
                doc.add_paragraph(paragraph.strip())
        
        doc.save(file_path)
    
    def _show_status_message(self, message: str, msg_type: str = "info"):
        
        color_map = {
            "success": "#28a745",
            "error": "#dc3545", 
            "warning": "#ffc107",
            "info": "#17a2b8"
        }
        
        color = color_map.get(msg_type, "#6c757d")
        self.status_label.setStyleSheet(f"color: {color}; font-weight: 600; font-size: 13px;")
        self.status_label.setText(message)
        
        if msg_type == "success":
            self.status_timer.start(5000)
    
    def _clear_status_message(self):
        
        self.status_label.clear()
        self.status_timer.stop()
    
    def closeEvent(self, event):
        
        if self.extract_thread and self.extract_thread.isRunning():
            self.extract_thread.stop()
        event.accept()

def main():
    
    app = QApplication(sys.argv)
    app.setApplicationName("PDF Text Extractor")
    app.setApplicationVersion("2.0")
    
    window = PDFExtractorGUI()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__": 
    main()