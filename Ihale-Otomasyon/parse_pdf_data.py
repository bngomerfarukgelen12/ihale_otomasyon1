import re

# Read the extracted text
with open('pdf_text.txt', 'r', encoding='utf-8') as f:
    text = f.read()

# Extract doors (AK, FK codes) and their quantities
doors_pattern = r'(AK\d|FK|FDK|MF|P\d|GC|K\d|ALCK|ALP|İŞLENEREK|CAMEKANLI|DEMİR KAPI|YANGIN KAPISI|MENFEZLİ YANGIN KAPISI)\s+([0-9/]+)?(?:.*?\n)*?\s*(\d+)\s*ADET'

matches = re.finditer(doors_pattern, text, re.IGNORECASE)

print("=" * 80)
print("DOORS AND WINDOWS EXTRACTED FROM PDF")
print("=" * 80)

doors_windows = []
for match in matches:
    code = match.group(1).strip()
    size = match.group(2) if match.group(2) else ""
    quantity = match.group(3)
    doors_windows.append({
        'code': code,
        'size': size,
        'quantity': quantity
    })
    print(f"{code:30} {size:20} {quantity:10} ADET")

print("\n" + "=" * 80)
print(f"Total items found: {len(doors_windows)}")

# Save to CSV for Excel
import csv
with open('doors_windows_extracted.csv', 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.DictWriter(f, fieldnames=['code', 'size', 'quantity'])
    writer.writeheader()
    writer.writerows(doors_windows)

print("Data saved to: doors_windows_extracted.csv")
