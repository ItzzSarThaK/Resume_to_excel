Resume Scanner - PDF to Excel Converter
A user-friendly desktop application that scans resumes (PDF/DOCX format) and extracts key information into a structured Excel file.

Features
📄 Multiple Format Support: Works with PDF and DOCX files
🔍 Smart Extraction: Automatically extracts:
Name
Email address
Phone number
Skills
Education details
Work experience
Summary/Objective
📊 Excel Export: Creates well-formatted Excel files with extracted data
🖥️ Easy-to-Use GUI: Simple graphical interface for seamless operation
🔧 OCR Support: Handles scanned PDFs using OCR technology
Installation
Prerequisites
Python 3.7 or higher
Tesseract OCR (for scanned PDFs)
Step 1: Install Python Dependencies
pip install -r requirements.txt
Step 2: Install Tesseract OCR (Optional, for scanned PDFs)
Windows:

Download Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki
Install it and note the installation path
Add Tesseract to your system PATH, or set it in the code:
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
macOS:

brew install tesseract
Linux:

sudo apt-get install tesseract-ocr
Usage
Run the application:

python resume_scanner.py
Select a resume:

Click "Select Resume (PDF/DOCX)" button
Choose your resume file
Scan and export:

Click "Scan & Export to Excel" button
Wait for the extraction process to complete
The Excel file will be saved in the same directory as your resume
Output Format
The Excel file contains the following sections:

Name: Extracted name from resume
Email: Email address
Phone: Phone number
Address: Physical address (if found)
Summary: Professional summary or objective
Skills: List of skills and competencies
Education: Educational qualifications
Experience: Work experience details
Requirements
pdfplumber==0.10.3
pytesseract==0.3.10
Pillow==10.2.0
openpyxl==3.1.2
pandas==2.2.0
python-docx==1.1.0
pdf2image==1.16.3
regex==2024.4.16
Notes
The application works best with text-based PDFs
For scanned PDFs, Tesseract OCR is required
The extraction accuracy depends on the resume format and structure
Some information might need manual verification
Troubleshooting
Issue: OCR not working

Make sure Tesseract is installed and added to PATH
For Windows, you may need to specify the Tesseract path in the code
Issue: Cannot extract text from PDF

The PDF might be password protected
The PDF might be corrupted
Try converting the PDF to a different format first
Issue: Missing information in Excel

The resume format might not be standard
Some information might be in images or non-standard formats
Manual review and editing may be required
License
This project is open source and available for personal and commercial use.
