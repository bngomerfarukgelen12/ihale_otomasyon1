import fitz  # pymupdf

# Extract text from PDF
doc = fitz.open('pdfler/MEB 16 DERSLİKLİ OKUL KAPI PENCERE DETAYLARI.pdf')
text = doc[0].get_text()

# Search for window entries (P1, P2, P3, etc.)
import re

# Looking for window patterns
patterns = {
    'P1': r'P1\s+(\d+)/(\d+).*?(?:ADET|adet)?',
    'P2': r'P2\s+(\d+)/(\d+).*?(?:ADET|adet)?',
    'P3': r'P3\s+(\d+)/(\d+).*?(?:ADET|adet)?',
    'MF': r'MF\s+(\d+)/(\d+).*?(?:ADET|adet)?',
    'FK': r'FK\s+(\d+)',
    'FDK': r'FDK\s*(\d+)\s+(\d+)/(\d+)',
}

print("SEARCHING FOR WINDOWS IN PDF:\n")

# Search for all "ADET" occurrences around windows
lines = text.split('\n')
window_data = []

for i, line in enumerate(lines):
    if any(win in line.upper() for win in ['P1', 'P2', 'P3', 'MF', 'FK', 'FDK', 'PENCERE', 'PVC']):
        print(f"Line {i}: {line}")
        # Check next few lines for ADET
        for j in range(1, 5):
            if i+j < len(lines) and 'ADET' in lines[i+j].upper():
                print(f"  -> {lines[i+j]}")
