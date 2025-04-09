import os
import re
import pytesseract
import pdfplumber
import docx
import pandas as pd
from PIL import Image
import shutil

# Folder paths
test_folder = r"D:\Studying\Python\invoiceFilter\invoiceFilter\test_emails"
processed_folder = r"D:\Studying\Python\invoiceFilter\invoiceFilter\processed"
manual_review_folder = r"D:\Studying\Python\invoiceFilter\invoiceFilter\manual_review"
enquiries_folder = r"D:\Studying\Python\invoiceFilter\invoiceFilter\enquiries"
likely_duplicates_folder = r"D:\Studying\Python\invoiceFilter\invoiceFilter\likely_duplicates_folder"
log_file = r"D:\Studying\Python\invoiceFilter\invoiceFilter\processed_invoices.txt"

# Create folders if they don't exist
os.makedirs(processed_folder, exist_ok=True)
os.makedirs(manual_review_folder, exist_ok=True)
os.makedirs(enquiries_folder, exist_ok=True)
os.makedirs(likely_duplicates_folder, exist_ok=True)

# Patterns
invoice_pattern = re.compile(r"(Invoice\s*No[:#]?\s*|Inv\s*[:#]?\s*|Inv\.?#?)\s*(\d+)", re.IGNORECASE)
abn_pattern = re.compile(r"\bABN[:\s]*([0-9 ]{11,20})\b", re.IGNORECASE)

def extract_invoice_number(text):
    match = invoice_pattern.search(text)
    if match:
        return match.group(2)
    fallback = re.search(r"\b\d{4,}\b", text)
    return fallback.group(0) if fallback else None

def extract_abn(text):
    match = abn_pattern.search(text)
    if match:
        return match.group(1)
    fallback = re.search(r"\b\d{11}\b", text)
    return fallback.group(0) if fallback else None

def extract_text_from_pdf(filepath):
    try:
        with pdfplumber.open(filepath) as pdf:
            return " ".join([page.extract_text() for page in pdf.pages if page.extract_text()])
    except Exception as e:
        print(f"❌ Error processing {os.path.basename(filepath)}: {e}")
        return ""

def extract_text_from_docx(filepath):
    try:
        doc = docx.Document(filepath)
        return " ".join([para.text for para in doc.paragraphs])
    except Exception as e:
        print(f"❌ Error processing {os.path.basename(filepath)}: {e}")
        return ""

def extract_text_from_excel(filepath):
    try:
        df = pd.read_excel(filepath, engine='openpyxl')
        return " ".join(df.astype(str).values.flatten())
    except Exception as e:
        print(f"❌ Error processing {os.path.basename(filepath)}: {e}")
        return ""

def extract_text_from_image(filepath):
    try:
        img = Image.open(filepath).convert('L')
        text = pytesseract.image_to_string(img)
        confidence_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)['conf']
        confidence = [int(c) for c in confidence_data if c.isdigit()]
        avg_confidence = sum(confidence) / len(confidence) if confidence else 0
        return text, avg_confidence
    except Exception as e:
        print(f"❌ Error processing {os.path.basename(filepath)}: {e}")
        return "", 0

def load_processed_records():
    if os.path.exists(log_file):
        with open(log_file, "r") as file:
            return set(file.read().splitlines())
    return set()

def save_processed_record(abn, invoice_number, filename):
    with open(log_file, "a") as file:
        file.write(f"{abn}|{invoice_number}|{filename}\n")

def move_file_safely(src, dest_folder):
    os.makedirs(dest_folder, exist_ok=True)
    base_name = os.path.basename(src)
    dest_path = os.path.join(dest_folder, base_name)

    counter = 1
    while os.path.exists(dest_path):
        name, ext = os.path.splitext(base_name)
        dest_path = os.path.join(dest_folder, f"{name} ({counter}){ext}")
        counter += 1

    shutil.move(src, dest_path)
    return dest_path

def process_test_files():
    processed_records = load_processed_records()

    for filename in os.listdir(test_folder):
        filepath = os.path.join(test_folder, filename)

        if not os.path.isfile(filepath):
            continue

        file_ext = os.path.splitext(filename)[1].lower()
        invoice_number = None
        abn = None
        low_confidence = False
        text = ""

        if file_ext in [".pdf"]:
            text = extract_text_from_pdf(filepath)
        elif file_ext in [".docx"]:
            text = extract_text_from_docx(filepath)
        elif file_ext in [".xls", ".xlsx"]:
            text = extract_text_from_excel(filepath)
        elif file_ext in [".png", ".jpg", ".jpeg"]:
            text, confidence = extract_text_from_image(filepath)
            if confidence < 60:
                low_confidence = True
        else:
            move_file_safely(filepath, enquiries_folder)
            print(f"📥 {filename} moved to Enquiries (unsupported format)")
            continue

        # Debug: show extracted text preview
        print(f"\n📄 File: {filename}")
        print(f"📝 Text Preview: {text[:500]}...")

        invoice_number = extract_invoice_number(text)
        abn = extract_abn(text)

        print(f"🔍 Invoice No: {invoice_number}")
        print(f"🔍 ABN: {abn}")

        if invoice_number and abn:
            record_key = f"{abn}|{invoice_number}|{filename}"
            duplicate_check_key = f"{abn}|{invoice_number}"

            if any(entry.startswith(duplicate_check_key + "|") for entry in processed_records):
                move_file_safely(filepath, likely_duplicates_folder)
                print(f"⚠️ Duplicate found for ABN {abn}, Invoice {invoice_number}. Moved to Likely Duplicates")
                continue

            save_processed_record(abn, invoice_number, filename)
            processed_records.add(record_key)

            if low_confidence:
                move_file_safely(filepath, manual_review_folder)
                print(f"📥 Moved to Manual Review (low OCR confidence)")
            else:
                move_file_safely(filepath, processed_folder)
                print(f"✅ Processed successfully")
        else:
            move_file_safely(filepath, manual_review_folder)
            print(f"📥 No invoice number or ABN found. Sent to Manual Review")

if __name__ == "__main__":
    process_test_files()

