import pandas as pd
from difflib import SequenceMatcher

excel_path = r'c:\Users\argge\OneDrive\Masaüstü\Ihale-Otomasyon\excel\csb insaat birim fiyatları.xlsx'

# Load Excel
df = pd.read_excel(excel_path, sheet_name=0, header=1, skiprows=0)
df.columns = [str(col).strip() for col in df.columns]
df = df.dropna(subset=['Poz No']).copy()

# Aranacak terimler
searches = [
    {
        'item': 'Row 23 - PVC su borusu',
        'keywords': ['PVC', 'boru', 'kanaleta'],
        'combined': 'PVC plastik boru su'
    },
    {
        'item': 'Row 26 - Hendek betonlaması',
        'keywords': ['hendek', 'beton kaplama', 'koruma'],
        'combined': 'hendek beton kaplan'
    }
]

def search_in_column(df, search_term, column='Tanım', min_ratio=0.25):
    results = []
    for idx, row in df.iterrows():
        tanım = row[column]
        if pd.isna(tanım):
            continue
            
        text = str(tanım).lower()
        search_lower = search_term.lower()
        
        # Exact keyword search
        if all(keyword.lower() in text for keyword in search_term.split()):
            ratio = 1.0
        else:
            ratio = SequenceMatcher(None, search_lower, text).ratio()
        
        if ratio > min_ratio:
            results.append({
                'ratio': ratio,
                'tanım': str(row['Tanım'])[:80],
                'birim': row['Birim'],
                'fiyat': row['Fiyat'],
                'poz': row['Poz No']
            })
    
    return sorted(results, key=lambda x: x['ratio'], reverse=True)[:3]

print('='*80)
print('ADVANCED SEARCH FOR MISSING ITEMS')
print('='*80)

for search in searches:
    print(f"\n{search['item']}")
    print("-"*80)
    
    results = search_in_column(df, search['combined'], min_ratio=0.2)
    
    if results:
        for i, r in enumerate(results, 1):
            print(f"{i}. [{r['ratio']:.0%}] Poz: {r['poz']}")
            print(f"   {r['tanım']}...")
            print(f"   Price: {r['fiyat']} {r['birim']}\n")
    else:
        print("  No matches found")
        # Try with just keywords
        print("\n  Trying with single keywords:")
        for keyword in search['keywords']:
            results = search_in_column(df, keyword, min_ratio=0.15)
            if results and results[0]['ratio'] > 0.3:
                print(f"  - Keyword '{keyword}': {results[0]['poz']} - {results[0]['fiyat']} {results[0]['birim']}")

print('\n' + '='*80)
