import openpyxl
import sys
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r'C:\Users\DTProject\Documents\Git\thetaleofher\thetaleofher_content_template.xlsx'
wb = openpyxl.load_workbook(excel_path, data_only=True)
ws = wb['Products']

names = []
slugs = []
rows_info = []

for r_idx, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    if r[0]:
        slug = str(r[0]).strip()
        name = str(r[1] or '').strip()
        cat = str(r[2] or '').strip()
        price = r[3] if r[3] else 0
        color = str(r[5] or '').strip()
        slugs.append(slug.lower())
        names.append(name)
        rows_info.append((r_idx, slug, name, cat, price, color))

print(f"Total baris produk terdaftar: {len(names)}\n")

name_counts = Counter([n.lower() for n in names])
slug_counts = Counter(slugs)

dup_names = {k: v for k, v in name_counts.items() if v > 1}
dup_slugs = {k: v for k, v in slug_counts.items() if v > 1}

print("=== 1. PENGECEKAN DUPLIKASI NAMA PRODUK (PRODUCT NAME) ===")
if dup_names:
    print(f"Ditemukan {len(dup_names)} nama produk duplikat:")
    for n, count in dup_names.items():
        print(f"  [DUPLIKAT] \"{n}\" muncul sebanyak {count} kali!")
        for r_idx, slug, name, cat, price, color in rows_info:
            if name.lower() == n:
                print(f"      - Baris {r_idx}: SKU={slug} | {name} | {cat} | Rp {price:,} | Warna: {color}")
else:
    print("-> TIDAK ADA nama produk yang sama persis (100% UNIK).")

print("\n=== 2. PENGECEKAN DUPLIKASI ID / SLUG (SKU) ===")
if dup_slugs:
    print(f"Ditemukan {len(dup_slugs)} ID/Slug duplikat:")
    for s, count in dup_slugs.items():
        print(f"  [DUPLIKAT] ID \"{s}\" muncul sebanyak {count} kali!")
else:
    print("-> TIDAK ADA ID/Slug produk yang sama (100% UNIK).")

print("\n=== 3. DAFTAR LENGKAP 16 PRODUK & VARIANNYA ===")
for idx, (r_idx, slug, name, cat, price, color) in enumerate(rows_info, 1):
    print(f"{idx:2d}. {name:<25} (SKU: {slug:<10} | Cat: {cat:<8} | Warna: {color})")
