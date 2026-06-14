import fitz  # pymupdf

doc = fitz.open('pdfler/MEB 16 DERSLİKLİ OKUL KAPI PENCERE DETAYLARI.pdf')
print(f"Total pages: {len(doc)}")

for page_num in range(len(doc)):
    page = doc[page_num]
    text = page.get_text()
    print(f"\n=== Page {page_num + 1} ===")
    print(text)
    print("\n" + "="*50)
