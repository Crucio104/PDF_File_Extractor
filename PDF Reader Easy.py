import PyPDF2
import re
from docx import Document
import customtkinter as ctk
import os
from tkinter import filedialog, messagebox

class PDFExctractorGUI:
     def __init__(self):
          self.window = ctk.CTk()
          self.window.title("PDF text Extractor")
          self.window.geometry("600x500")

          self.filepath = ""
          self.extractedtext = ""

          self.create_widgets()
     
     def create_widgets(self):
            title_lable = ctk.CTkLabel(self.window, text = "PDF Text Extractor", font = ctk.CTkFont(size = 20, weight = "bold"))
            title_lable.pack(pady=20)
            

def text_extractor(path_pdf : str):
    text = ""
    try:
        with open(path_pdf, "rb") as file:
            reader =  PyPDF2.PdfReader(file)
            for page in range(len(reader.pages)):
                page_text = reader.pages[page].extract_text()
                text += page_text + "\n"
        return text       
    except Exception as e:
            print(f"Error during the extracion: {e}")
            return None
    
def text_saver_txt(text : str) -> None:
     with open("/home/crucio/Desktop/File python/Extracted_text.txt", "w", encoding = "utf-8") as file:
          file.write(text)

def text_saver_docx(text : str) -> None:
    doc = Document()
    doc.add_paragraph(text)
    doc.save("/home/crucio/Desktop/File python/Extracted_text.docx")

def text_cleaner(text : str) -> str:
        text = re.sub(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]', '', text)
        text = text.replace('\x00', '')  # Byte NULL
        text = text.replace('\uffff', '')  # Unicode replacement character
        return text
          

if __name__ == "__main__":
     path_pdf = "/home/crucio/Downloads/Calc unit 1 January.pdf"
     text = text_extractor(path_pdf)
     text = text_cleaner(text)
     try:
        text_saver_txt(text)
        text_saver_docx(text)
     except:
          print("Couldn't save the file.")
     print(text)
