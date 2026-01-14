import os
import pytesseract
import pdfplumber
import requests
import openpyxl
from PIL import Image

# ---------------- TESSERACT PATH (PORTABLE) ----------------
# Set this once in Windows:
# setx TESSERACT_PATH "C:\Users\Tushar Gupta\Desktop\MY_pytesseract\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = os.getenv("TESSERACT_PATH")

# ---------------- IMAGE EXTRACTION ----------------
def extract_from_image(file_path):
    img = Image.open(file_path)
    return pytesseract.image_to_string(img)


# ---------------- PDF EXTRACTION (NO OCR FALLBACK HERE YET) ----------------
def extract_from_pdf(file_path):
    text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


# ---------------- MAIN EXTRACTOR ----------------
def extract_text(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"]:
        return extract_from_image(file_path)

    elif ext == ".pdf":
        return extract_from_pdf(file_path)

    else:
        return "Unsupported file type."


# ---------------- LLM LOGIC WITH UPDATED STRICT PROMPT ----------------
def extract_salary_info_text(extracted_text, requested_doc_type, requested_period):
    prompt = f"""
Your task is to extract EXACTLY four fields using STRICT rule-based logic. No guessing or inference.

OUTPUT FORMAT (strict):
Document Type: ...
Month/Year: ...
Validation: ...
Estimated Monthly Salary: ...

----------------------------------------------------------
1. DOCUMENT TYPE DETECTION (STRICT)
----------------------------------------------------------

Detect the Document type only from explicit keywords.

Payslip indicators (any → Payslip):
- Payslip / Salary Slip
- Pay Period
- Basic Pay
- Incentive Pay
- House Rent Allowance
- Meal Allowance
- Total Earnings / Total Deductions
- Net Pay / Total Amount / Take Home
- Employer Signature / Employee Signature

IT Return indicators (any → IT Return):
- INDIAN INCOME TAX RETURN ACKNOWLEDGEMENT
- Income Tax Return
- ITR-1 / ITR-2 / ITR-3 / ITR-4 / ITR-5 / ITR-6 / ITR-7
- Assessment Year / AY 20
- Financial Year / FY 20
- Filed u/s
- Total Income / Gross Total Income
- ITR Verification Form
- Verification Code

If none match → Document type = Invalid Document.

----------------------------------------------------------
2. REQUESTED TYPE VALIDATION (STRICT)
----------------------------------------------------------

Requested type = {requested_doc_type}

If Document type ≠ requested type:
  Document Type = Document type
  Validation = Fail - Document type mismatch
  Month/Year = Not Applicable
  Estimated Monthly Salary = Not Applicable
  STOP and output only the four lines.

----------------------------------------------------------
3. PERIOD EXTRACTION (STRICT)
----------------------------------------------------------

Requested Period = {requested_period}

IT Return RULES → Extract AY or FY periods ONLY:  
       e.g., "AY 2024-25", "FY 2024-25", "2024-2025", "2024 to 2025"
   - Special rule:
        Compare ENTIRE period (example: "2024-2025" MUST equal "2024-2025")
        "2024-2025" ≠ "2025-2026"
        Matching only the last 2 digits is NOT allowed unless the entire range matches.
  
PAYSLIP RULES → Extract explicit Month + Year only:
        E.g., "May 2024", "06/2023", "August 2025"
   - Month and Year must both match EXACTLY to requested_period.

Rules:
- If no valid period found → Validation = Fail - Period not found
- Else if extracted period ≠ requested period → Validation = Fail - Period mismatch
- Else if →if Document type = requested type Validation = Pass
           else Validation = Fail - Document type mismatch

Month/Year = extracted period even if Validation = Fail

----------------------------------------------------------
4. SALARY EXTRACTION (STRICT)
----------------------------------------------------------

If Validation ≠ Pass:
  Estimated Monthly Salary = Not Applicable

If Validation = Pass:
  IT Return → Extract “Total Income” or “Gross Total Income” and divide by 12.
  Payslip → Extract “Net Pay” (preferred) or “Gross Pay”.

----------------------------------------------------------
STRICT OUTPUT RULES
----------------------------------------------------------
- Output ONLY the four lines.
- No explanation.
- No extra text.
----------------------------------------------------------
DOCUMENT TEXT:
{extracted_text}
"""




    response = requests.post(
        "http://localhost:11434/api/generate",
        json={"model": "qwen2.5:7b-instruct", "prompt": prompt, "stream": False}
    ).json()

    if "response" in response:
        return response["response"]
    else:
        raise Exception(f"Ollama error: {response}")


# ---------------- SAVE TO EXCEL ----------------
def save_to_excel_append(output_text, file_path, excel_file="summary.xlsx"):

    if os.path.exists(excel_file):
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Summary"
        ws.append([
            "Filename",
            "Document Type",
            "Month/Year",
            "Validation",
            "Estimated Monthly Salary"
        ])

    headers = [
        "Filename",
        "Document Type",
        "Month/Year",
        "Validation",
        "Estimated Monthly Salary"
    ]

    mapping = {h: "" for h in headers}
    mapping["Filename"] = os.path.basename(file_path)

    for line in output_text.split("\n"):
        if ":" in line:
            key, value = line.split(":", 1)
            key = key.strip().lower()
            value = value.strip()

            if "document type" in key:
                mapping["Document Type"] = value
            elif "month/year" in key:
                mapping["Month/Year"] = value
            elif "validation" in key:
                mapping["Validation"] = value
            elif "estimated monthly salary" in key:
                mapping["Estimated Monthly Salary"] = value

    ws.append([mapping[h] for h in headers])
    wb.save(excel_file)
    print(f"Row added for {file_path}")


# ---------------- EXECUTION LOOP ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
documents_folder = os.path.join(BASE_DIR, "documents")
requested_doc_type = "IT Return" #"Payslip"       # <-- YOU MUST SET THIS
requested_period = "2024-25" #"May 2025"       # <-- AND THIS ("May 2025" for payslip)

for filename in os.listdir(documents_folder):
    file_path = os.path.join(documents_folder, filename)

    if not os.path.isfile(file_path):
        continue

    print(f"\nProcessing: {filename}")

    extracted_text = extract_text(file_path)

    plain_output = extract_salary_info_text(
        extracted_text,
        requested_doc_type,
        requested_period
    )

    print("LLM Output:\n", plain_output)

    save_to_excel_append(plain_output, file_path)
