import pandas as pd
import PyPDF2
import re

# CSB İnşaat PDF'sinden arayacağımız terimler
search_terms = [
    "PVC plastik pis su borusu",
    "PVC pipe",
    "drenaj borusu",
    "hendek",
    "beton kaplama",
    "beton lining"
]

# CSB İnşaat PDF dosyası
pdf_path = r'c:\Users\argge\OneDrive\Masaüstü\Ihale-Otomasyon\pdfler\CSB İNŞAAT BİRİM FİYAT LİSTESİ.pdf'

print("PDF'den aranacak kalemler:")
print("1. PVC plastik pis su borusu (çap: 100-110 mm)")
print("2. Hendeklerin betonla kaplanması (C30/37)")
print()

try:
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        print(f"PDF toplam {len(pdf_reader.pages)} sayfadan oluşuyor")
        
        # Aranacak terimleri PDF'de ara
        found_items = []
        
        for page_num in range(min(50, len(pdf_reader.pages))):  # İlk 50 sayfayı ara
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            
            for term in search_terms:
                if term.lower() in text.lower():
                    found_items.append({
                        'page': page_num + 1,
                        'term': term,
                        'text_snippet': text[max(0, text.lower().find(term.lower()) - 100):min(len(text), text.lower().find(term.lower()) + 200)]
                    })
        
        if found_items:
            print(f"\n{len(found_items)} sonuç bulundu:")
            for item in found_items[:10]:
                print(f"  Sayfa {item['page']}: {item['term']}")
                print(f"    -> {item['text_snippet'][:150]}")
        else:
            print("\nAranılan terimler PDF'de bulunamadı")
            
except Exception as e:
    print(f"Hata: {e}")
