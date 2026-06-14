import pdfplumber
import json

pdf = pdfplumber.open('pdfler/MEB 16 DERSLİKLİ OKUL KAPI PENCERE DETAYLARI.pdf')
page = pdf.pages[0]

# Try to extract tables
tables = page.extract_tables()
if tables:
    print("Found tables:")
    for i, table in enumerate(tables):
        print(f"\n--- Table {i+1} ---")
        for row in table:
            print(row)
else:
    print("No tables found. Extracting text:")
    print(page.extract_text())
