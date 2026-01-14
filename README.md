# Automated Document Intelligence – OCR + LLM

An intelligent document processing system that extracts, validates, and computes salary information from Payslips and Income Tax Returns (ITR) using OCR and a local LLM.

This project simulates how financial institutions validate income documents for loans, KYC, and payroll processing.

---

## Features

- OCR for scanned images using Tesseract  
- Text extraction from PDFs using pdfplumber  
- LLM-based document classification (Payslip vs IT Return)  
- Strict period validation (Month/Year or AY/FY)  
- Salary extraction using Qwen 2.5 via Ollama  
- Automatic Excel report generation  

---

## Project Structure

```text
Automated_Document_Intelligence_OCR_LLM
├── document_info_extraction.py
├── requirements.txt
├── README.md
└── documents
    ├── ITRV-1_r.pdf
    └── payslip.jpg
```
---

## Prerequisites

### Install Tesseract OCR
Download from:  
https://github.com/UB-Mannheim/tesseract/wiki  

After installation, note the path to `tesseract.exe`.

---

## Environment Variables

This project uses an environment variable to locate Tesseract.

On Windows (run once in Command Prompt):

setx TESSERACT_PATH "C:\Program Files\Tesseract-OCR\tesseract.exe"


Restart VS Code after running this command.

The Python code reads this value using `os.getenv("TESSERACT_PATH")` so the project works on any machine.

---

## Install Dependencies

pip install -r requirements.txt


---

## Local LLM Setup (Ollama)

Install Ollama from:  
https://ollama.com  

Then run:

ollama run qwen2.5:7b-instruct


This starts the local LLM server on port 11434.

---

## How to Run

1. Place all PDFs and images inside the `documents` folder  
2. Set the requested document type and period inside `document_info_extraction.py`

Example for IT Return:

requested_doc_type = "IT Return"
requested_period = "2024-25"


Example for Payslip:

requested_doc_type = "Payslip"
requested_period = "May 2025"


3. Run:

python document_info_extraction.py


4. Output will be written to:

summary.xlsx


---

## Validation Logic

The system validates uploaded documents against user-requested inputs.

| Condition | Result |
|--------|--------|
Document type mismatch | Fail |
Period mismatch | Fail |
Both match | Salary is extracted |

This prevents incorrect or fraudulent document submissions.

---

## Output Format

The Excel file `summary.xlsx` contains:

| Filename | Document Type | Month/Year | Validation | Estimated Monthly Salary |

---

## Use Cases

- Loan application income verification  
- Payroll and HR document validation  
- KYC and FinTech automation  
- Financial document intelligence systems  

---

## Tech Stack

- Python  
- Tesseract OCR  
- pdfplumber  
- Ollama (Qwen 2.5)  
- OpenPyXL  