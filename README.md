# invoiceFilter

This is a basic invoice filtering system designed originally for a outlook inbox, however it is currently setup to work on a local directory for testing purposes.

How it works:

1. Identifies file types, and if the email has an attachement
2. If the file does not have either an attachement, it will be treated as an inquiry. If there is an attachement that is not supported, it will send it for manual review. Supported files continue to step 3.
3. Scans the attachements for invoice numbers, ABN and saves the filename searching for commonly used formats sure as "INV# 02003" or "Invoice: 1002"
4. If an invoice number is found, it will check against saved invoice numbers for duplicates
5. If the script is not confident, it will be sent for manual review
6. If the script is confident + invoice number is not a duplicate, invoice will be sent to processing folder (to be replaced by automatically sending to CRM once highly confident + in actual use)

Libraries used:
1. os
2. re
3. pytesseract
4. pdfplumber
5. docx
6. pandas
7. Image from PIL 
8. shutil

To be implemented:

1. ABN and filename matching for duplicates
2. Outlook implementation - needs access to a clone of the inbox/small amount to trial with

Future improvements:

Once this system is deployed, future improvements could include the following. 

1. Save email address with invoice number to avoid mislabelling common invoice numbers from different sources as duplicates.
2. Use machine learning by feeding a reviewed set of invoices to increase accuracy and reduce human review element
3. 
