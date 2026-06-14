import pandas as pd

excel_path = r'c:\Users\argge\OneDrive\Masaüstü\Ihale-Otomasyon\excel\csb insaat birim fiyatları.xlsx'

# Row 2'den başla (skiprows=2)
df = pd.read_excel(excel_path, sheet_name=0, skiprows=2)

print('Kolon adları:')
for col in df.columns:
    print(f'  "{col}"')

print('\nİlk 3 satır:')
print(df.head(3))
