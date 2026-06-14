import pandas as pd

# Mevcut fiyat dosyalarını kontrol et
print('=== csb_poz_numaralari.xlsx ===')
poz_nos = pd.read_excel(r'c:\Users\argge\OneDrive\Masaüstü\Ihale-Otomasyon\csb_poz_numaralari.xlsx')
print(poz_nos[['Sıra', 'Poz No', 'Fiyat (TL)']].head(10))
print(f"Fiyat sütunu doldurulmuş: {poz_nos['Fiyat (TL)'].notna().sum()} / {len(poz_nos)}")

print('\n=== csb insaat birim fiyatları.xlsx ===')
insaat = pd.read_excel(r'c:\Users\argge\OneDrive\Masaüstü\Ihale-Otomasyon\excel\csb insaat birim fiyatları.xlsx')
print(f'Sütunlar: {insaat.columns.tolist()}')
print(insaat.head(5))

print('\n=== csb_tum_fiyatlar.xlsx ===')
tum = pd.read_excel(r'c:\Users\argge\OneDrive\Masaüstü\Ihale-Otomasyon\excel\csb_tum_fiyatlar.xlsx')
print(f'Sütunlar: {tum.columns.tolist()}')
print(tum.head(5))
