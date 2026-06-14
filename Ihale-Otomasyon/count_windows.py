import fitz
import re

doc = fitz.open('pdfler/MEB 16 DERSLİKLİ OKUL KAPI PENCERE DETAYLARI.pdf')
text = doc[0].get_text()

# Count occurrences of each window type
windows = {
    'P1 240/200': 0,
    'P2 200/200': 0,
    'P3 80/60': 0,
    'FDK1 360/280': 0,
    'FK': 0,
    'MF 60/60': 0,
}

# Count exact matches
windows['P1 240/200'] = len(re.findall(r'P1\s+240/?200', text, re.IGNORECASE))
windows['P2 200/200'] = len(re.findall(r'P2\s+200/?200', text, re.IGNORECASE))
windows['P3 80/60'] = len(re.findall(r'P3\s+80/?60', text, re.IGNORECASE))
windows['FDK1 360/280'] = len(re.findall(r'FDK1?\s+360/?280', text, re.IGNORECASE))
windows['FK'] = len(re.findall(r'\bFK\b', text, re.IGNORECASE))
windows['MF 60/60'] = len(re.findall(r'MF\s+60/?60', text, re.IGNORECASE))

print("=" * 60)
print("PENCERE SAYILARI (WINDOW QUANTITIES)")
print("=" * 60)
print()

total = 0
for window, count in windows.items():
    print(f"{window:20} : {count:3} ADET")
    total += count

print()
print("=" * 60)
print(f"TOPLAM PENCERE SAYISI: {total} ADET")
print("=" * 60)
