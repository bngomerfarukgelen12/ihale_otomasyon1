import pandas as pd
import fitz  # pymupdf
import re

# Eşleşmesi olmayan kalemler
missing_items = [
    {
        'row': 23,
        'description': 'Sert PVC plastik pis su borusu (geçme muflu, çap: 100-110 mm)',
        'keywords': ['PVC', 'pis su borusu', 'drenaj', 'boru', '100', '110', 'mm']
    },
    {
        'row': 26,
        'description': 'Hendeklerin betonla kaplanması (Orta refüj ve yarma hendeği) (C30/37)',
        'keywords': ['hendek', 'beton kaplama', 'beton', 'C30/37', 'refüj']
    }
]

print("="*80)
print("Missing Items - PDF Search")
print("="*80)
print()

# First, try Excel file with better fuzzy matching
excel_path = r'c:\Users\argge\OneDrive\Masaüstü\Ihale-Otomasyon\excel\csb insaat birim fiyatları.xlsx'

# Load Excel
df = pd.read_excel(excel_path, sheet_name=0, header=1, skiprows=0)
df.columns = [str(col).strip() for col in df.columns]
df = df.dropna(subset=['Poz No']).copy()

from difflib import SequenceMatcher

def find_similar(search_term, df, threshold=0.3):
    results = []
    for idx, row in df.iterrows():
        tanım = str(row['Tanım'])
        ratio = SequenceMatcher(None, search_term.lower(), tanım.lower()).ratio()
        if ratio > threshold:
            results.append({
                'ratio': ratio,
                'tanım': tanım[:100],
                'birim': row['Birim'],
                'fiyat': row['Fiyat'],
                'poz': row['Poz No']
            })
    
    return sorted(results, key=lambda x: x['ratio'], reverse=True)[:5]

for item in missing_items:
    print(f"\n[Row {item['row']}] {item['description']}")
    print("-" * 80)
    
    # Search with keywords
    for keyword in item['keywords']:
        results = find_similar(keyword, df, threshold=0.2)
        if results:
            print(f"\n  Keyword '{keyword}' - Top matches:")
            for r in results[:2]:
                print(f"    [{r['ratio']:.0%}] {r['tanım']} - {r['fiyat']} {r['birim']} (Poz: {r['poz']})")
            break

print("\n" + "="*80)
