# PDF Converter for Pashto and Dari

This application converts PDF files to Word documents with optional OCR support for Pashto and Dari text.

## Features

- View PDF files with zoom controls.
- Convert using Microsoft Word (requires Windows and Word installed).
- Convert using Python libraries with optional Tesseract OCR.
- Supports interface languages: English, Dari and Pashto.

## Requirements

- Python 3.7+
- PyQt5
- pytesseract
- pdf2image
- pdf2docx
- PyMuPDF
- Pillow
- docx

OCR requires Tesseract with Pashto and Dari language data.

## Running

From the project directory:

```bash
python app.py
```

Use the GUI to open a PDF, set conversion options and save the DOCX file.
