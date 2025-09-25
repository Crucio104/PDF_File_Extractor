FROM python:3.13

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
RUN apt-get update && apt-get install -y tesseract-ocr tesseract-ocr-eng

COPY . .
CMD ["python", "pdf_extractor.py"]