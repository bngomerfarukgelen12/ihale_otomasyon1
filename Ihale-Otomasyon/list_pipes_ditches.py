import pandas as pd

excel_path = r'c:\Users\argge\OneDrive\Masaüstü\Ihale-Otomasyon\excel\csb insaat birim fiyatları.xlsx'

# Load Excel
df = pd.read_excel(excel_path, sheet_name=0, header=1, skiprows=0)
df.columns = [str(col).strip() for col in df.columns]
df = df.dropna(subset=['Poz No']).copy()

print('='*100)
print('BORULAR (PIPES) - All items containing "boru"')
print('='*100)

boru_items = df[df['Tanım'].astype(str).str.contains('boru|BORU', case=False, na=False)]
for idx, row in boru_items.iterrows():
    print(f"\nPoz: {row['Poz No']} | Birim: {row['Birim']}")
    print(f"  {row['Tanım']}")
    print(f"  Fiyat: {row['Fiyat']}")

print('\n' + '='*100)
print('HENDEKKLER (DITCHES) - All items containing "hendek"')
print('='*100)

hendek_items = df[df['Tanım'].astype(str).str.contains('hendek|HENDEK|drenaj|DRENAJ|kanalet', case=False, na=False)]
for idx, row in hendek_items.iterrows():
    print(f"\nPoz: {row['Poz No']} | Birim: {row['Birim']}")
    print(f"  {row['Tanım']}")
    print(f"  Fiyat: {row['Fiyat']}")

print('\n' + '='*100)
print(f"Total items with 'boru': {len(boru_items)}")
print(f"Total items with 'hendek': {len(hendek_items)}")
