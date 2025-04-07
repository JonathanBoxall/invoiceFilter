import os
import re
import pytesseract
import pdfplumber
import docx
import pandas as pd
from PIL import Image
import shutil

# Folder paths
test_folder = r"D:\Studying\Python\invoice_test_env\test_emails"
processed_folder = r"D:\Studying\Python\invoice_test_env\processed"
manual_review_folder = r"D:\Studying\Python\invoice_test_env\manual_review"
enquiries_folder = r"D:\Studying\Python\invoice_test_env\enquiries"
log_file = r"D:\Studying\Python\invoice_test_env\processed_invoices.txt"

# Create folders if they don't exist
os.makedirs(processed_folder, exist_ok=True)
os.makedirs(manual_review_folder, exist_ok=True)
os.makedirs(enquiries_folder, exist_ok=True)

# Improved invoice number regex
invoice_pattern = re.compile(
    r"(Invoice\s*No\.?\s*[:\-]?\s*|Inv\s*#?\s*[:\-]?\s*)\s*(\d+)",
    re.IGNORECASE
)

def extract_invoice_number(text):
    cleaned = text.replace('\xa0', ' ').replace('\n', ' ').strip()
    match = invoice_pattern.search(cleaned)
    if match:
        print(f"[DEBUG] Match found: '{match.group(0)}' -> Invoice Number: {match.group(2)}")
    else:
        print(f"[DEBUG] No invoice number match found in:\n{cleaned[:300]}...\n")
    return match.group(2) if match else None

def extract_text_from_pdf(filepath):
    with pdfplumber.open(filepath) as pdf:
        return " ".join([page.extract_text() for page in pdf.pages if page.extract_text()])

def extract_text_from_docx(filepath):
    doc = docx.Document(filepath)
    text = " ".join([para.text for para in doc.paragraphs])
    print(f"[DEBUG] Extracted DOCX Text from {os.path.basename(filepath)}:\n{text}\n")
    return text

def extract_text_from_excel(filepath):
    df = pd.read_excel(filepath, engine='openpyxl')
    return " ".join(df.astype(str).values.flatten())

def extract_text_from_image(filepath):
    img = Image.open(filepath).convert('L')  # Convert to grayscale
    text = pytesseract.image_to_string(img)
    confidence_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)['conf']
    confidence = [int(c) for c in confidence_data if c.isdigit()]
    avg_confidence = sum(confidence) / len(confidence) if confidence else 0
    return text, avg_confidence

def load_processed_invoices():
    if os.path.exists(log_file):
        with open(log_file, "r") as file:
            return set(file.read().splitlines())
    return set()

def save_processed_invoice(invoice_number):
    with open(log_file, "a") as file:
        file.write(f"{invoice_number}\n")

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
    processed_invoices = load_processed_invoices()

    for filename in os.listdir(test_folder):
        filepath = os.path.join(test_folder, filename)

        if not os.path.isfile(filepath):
            continue

        file_ext = os.path.splitext(filename)[1].lower()
        invoice_number = None
        low_confidence = False

        print(f"\n[INFO] Processing {filename}")

        try:
            if file_ext == ".pdf":
                text = extract_text_from_pdf(filepath)
                invoice_number = extract_invoice_number(text)

            elif file_ext == ".docx":
                text = extract_text_from_docx(filepath)
                invoice_number = extract_invoice_number(text)

            elif file_ext in [".xls", ".xlsx"]:
                text = extract_text_from_excel(filepath)
                invoice_number = extract_invoice_number(text)

            elif file_ext in [".png", ".jpg", ".jpeg"]:
                text, confidence = extract_text_from_image(filepath)
                invoice_number = extract_invoice_number(text)
                if confidence < 60:
                    low_confidence = True

            else:
                # Unrecognized file extension (likely enquiry or unsupported)
                move_file_safely(filepath, enquiries_folder)
                print(f"[INFO] Moved to Enquiries (Unknown Type): {filename}")
                continue

            if invoice_number:
                if invoice_number in processed_invoices:
                    print(f"[INFO] Skipping duplicate invoice number: {invoice_number} ({filename})")
                    continue

                processed_invoices.add(invoice_number)
                save_processed_invoice(invoice_number)

                if low_confidence:
                    move_file_safely(filepath, manual_review_folder)
                    print(f"[INFO] Moved to Manual Review (Low OCR Confidence): {filename} (Invoice No: {invoice_number})")
                else:
                    move_file_safely(filepath, processed_folder)
                    print(f"[INFO] Successfully Processed: {filename} (Invoice No: {invoice_number})")
            else:
                move_file_safely(filepath, manual_review_folder)
                print(f"[INFO] Moved to Manual Review (No Invoice No): {filename}")

        except Exception as e:
            print(f"[ERROR] Failed to process {filename}: {e}")
            move_file_safely(filepath, manual_review_folder)

if __name__ == "__main__":
    process_test_files()
