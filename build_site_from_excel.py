import openpyxl
import json
import re
import os

EXCEL_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\thetaleofher_content_template.xlsx'
INDEX_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\index.html'
KATALOG_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\katalog.html'
PRODUK_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\produk.html'
SITEMAP_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\sitemap.xml'

COLOR_HEX_MAP = {
    'light blue': '#ADC7D9',
    'ivory': '#F5EFE6',
    'army': '#555E4C',
    'white': '#FAF6F1',
    'black': '#241A1F',
    'cream olive': '#C2BAA6',
    'taupe': '#B49C8A',
    'navy': '#2B2F3A'
}

def get_color_hex(color_name):
    clean_name = color_name.strip().lower()
    for k, v in COLOR_HEX_MAP.items():
        if k in clean_name:
            return v
    return '#E5DCD3'

def build():
    print("Loading Excel template...")
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    # 1. Read General Info
    gen_info = {}
    ws_gen = wb['General_Info']
    for r in ws_gen.iter_rows(min_row=2, values_only=True):
        if r[0]:
            gen_info[str(r[0]).strip()] = str(r[2] or '').strip()

    # 2. Read Products
    ws_prod = wb['Products']
    products = []
    for r in ws_prod.iter_rows(min_row=2, values_only=True):
        if r[0]:
            pid = str(r[0]).strip().lower()
            name = str(r[1] or '').strip()
            cat = str(r[2] or '').strip().lower()
            price = int(r[3]) if r[3] else 0
            tag = str(r[4] or '').strip()
            color_str = str(r[5] or '').strip()
            sizing = str(r[6] or 'One Size').strip()
            material = str(r[7] or '').strip()
            fit = str(r[8] or '').strip()
            desc = str(r[9] or '').strip()
            care = str(r[10] or '').strip()
            front_img = str(r[11] or '').strip().lower()
            back_img = str(r[12] or '').strip().lower()
            gallery_str = str(r[13] or '').strip().lower()
            gallery = [g.strip() for g in gallery_str.split(',') if g.strip()]
            status = str(r[14] or 'Tampil').strip()
            
            if not gallery and front_img:
                gallery = [front_img]
            if back_img and back_img not in gallery:
                gallery.insert(1, back_img)
                
            color_name = color_str if color_str else 'Standard'
            hex_code = get_color_hex(color_name)

            products.append({
                'id': pid,
                'name': name,
                'cat': cat,
                'price': price,
                'tag': tag,
                'color': color_name,
                'hex': hex_code,
                'sizing': sizing,
                'material': material,
                'fit': fit,
                'desc': desc,
                'care': care,
                'front_img': front_img,
                'back_img': back_img if back_img else front_img,
                'gallery': gallery,
                'status': status
            })

    print(f"Loaded {len(products)} products from Excel.")

    # -------------------------------------------------------------
    # A. UPDATE INDEX.HTML
    # -------------------------------------------------------------
    print("Updating index.html...")
    with open(INDEX_FILE, 'r', encoding='utf-8') as f:
        idx_content = f.read()

    # Build 4 featured products for homepage
    featured = products[:4]
    feat_html_list = []
    delays = ['', 'd1', 'd2', 'd3']
    for i, p in enumerate(featured):
        d_class = delays[i] if i < len(delays) else ''
        feat_card = f'''      <a href="produk.html?id={p['id']}&c=0" class="prod reveal {d_class} group block">
        <div class="relative overflow-hidden aspect-[3/4] bg-[var(--paper-2)]">
          <img src="{p['front_img']}" alt="{p['name']} {p['color']}" class="w-full h-full object-cover object-top">
          <span class="quick absolute bottom-3 left-1/2 -translate-x-1/2 bg-[var(--paper)]/95 text-[var(--ink)] text-[10px] t2 uppercase px-5 py-2.5 whitespace-nowrap">View Details</span>
        </div>
        <div class="flex items-start justify-between mt-4">
          <div><h4 class="f text-lg leading-tight">{p['name']}</h4><p class="text-[12px] text-[var(--stone)] mt-0.5">{p['color']}</p></div>
          <p class="text-[13px] whitespace-nowrap">Rp {p['price']:,}</p>
        </div>
      </a>'''
        feat_html_list.append(feat_card)
    feat_grid_html = '\n\n'.join(feat_html_list)

    # Custom replacement using regex find
    match = re.search(r'(<div class="grid grid-cols-2 lg:grid-cols-4 gap-x-5 gap-y-14">)(.*?)(</div>\s*</section>)', idx_content, re.DOTALL)
    if match:
        idx_content = idx_content[:match.start(2)] + '\n' + feat_grid_html + '\n    ' + idx_content[match.end(2):]

    # Update editorial band image to first product
    idx_content = re.sub(r'assets/vittoria-black-1\.jpg', products[0]['front_img'], idx_content)
    idx_content = re.sub(r'assets/camalia-taupe-6\.jpg', products[1]['front_img'], idx_content)

    with open(INDEX_FILE, 'w', encoding='utf-8') as f:
        f.write(idx_content)
    print("[OK] index.html updated.")

    # -------------------------------------------------------------
    # B. UPDATE KATALOG.HTML
    # -------------------------------------------------------------
    print("Updating katalog.html...")
    with open(KATALOG_FILE, 'r', encoding='utf-8') as f:
        kat_content = f.read()

    # Collect distinct categories
    cats = []
    for p in products:
        c_norm = p['cat'].lower()
        if c_norm not in cats:
            cats.append(c_norm)

    cat_chips = ['<button class="chip active" data-cat="all">All</button>']
    for c in cats:
        cat_chips.append(f'<button class="chip" data-cat="{c}">{c.title()}</button>')
    cat_chips_html = ''.join(cat_chips)

    # Replace category chips in toolbar
    kat_content = re.sub(
        r'<div class="flex items-center gap-2\.5 overflow-x-auto no-scrollbar">.*?</div>',
        lambda m: f'<div class="flex items-center gap-2.5 overflow-x-auto no-scrollbar">{cat_chips_html}</div>',
        kat_content
    )

    # Build products JS array for katalog.html
    js_prods = []
    for p in products:
        tag_prop = f"tag:'{p['tag']}', " if p['tag'] else ""
        js_prods.append(
            f"    {{name:'{p['name']}', color:'{p['color']}', price:{p['price']}, cat:'{p['cat']}', {tag_prop}fu:'{p['front_img']}', bu:'{p['back_img']}', href:'produk.html?id={p['id']}&c=0', sw:['{p['hex']}'], imgs:['{p['front_img']}']}}"
        )
    js_prods_str = 'const products = [\n' + ',\n'.join(js_prods) + '\n  ];'

    # Replace products array in katalog.html
    kat_content = re.sub(
        r'const products = \[\s*\{.*?\}\s*\];',
        lambda m: js_prods_str,
        kat_content,
        flags=re.DOTALL
    )

    with open(KATALOG_FILE, 'w', encoding='utf-8') as f:
        f.write(kat_content)
    print("[OK] katalog.html updated.")

    # -------------------------------------------------------------
    # C. UPDATE PRODUK.HTML
    # -------------------------------------------------------------
    print("Updating produk.html...")
    with open(PRODUK_FILE, 'r', encoding='utf-8') as f:
        prd_content = f.read()

    # Build PRODUCTS dictionary for produk.html
    prod_dict_entries = []
    related_entries = []

    for p in products:
        shots_json = json.dumps(p['gallery'])
        color_obj = f"{{name:'{p['color']}', hex:'{p['hex']}', img:'{p['front_img']}', shots:{shots_json}}}"
        mat_text = p['material']
        if p['care']:
            mat_text += f" {p['care']}"
            
        entry = f"""    {p['id']}: {{
      name: {json.dumps(p['name'])},
      price: {p['price']},
      line: 'The Tale of Her · The Summer Capsule',
      desc: {json.dumps(p['desc'])},
      fit: {json.dumps(p['fit'])},
      mat: {json.dumps(mat_text)},
      ship: 'Shipping: Orders placed before 04:00 p.m. will be processed on the same day.\\n\\nReturn: Exchanges accepted within 3 days with unboxing video via WhatsApp +62-818-0505-2929.',
      shopeeLink: '',
      colors: [ {color_obj} ]
    }}"""
        prod_dict_entries.append(entry)
        related_entries.append(
            f"    {{id:'{p['id']}', c:0, name:'{p['name']}', color:'{p['color']}', price:{p['price']}, img:'{p['front_img']}'}}"
        )

    products_obj_str = "const PRODUCTS = {\n" + ",\n".join(prod_dict_entries) + "\n  };"
    related_obj_str = "const RELATED = [\n" + ",\n".join(related_entries) + "\n  ];"

    # Replace PRODUCTS in produk.html
    prd_content = re.sub(
        r'const PRODUCTS = \{.*?\n  \};',
        lambda m: products_obj_str,
        prd_content,
        flags=re.DOTALL
    )

    # Replace RELATED in produk.html
    prd_content = re.sub(
        r'const RELATED = \[.*?\n  \];',
        lambda m: related_obj_str,
        prd_content,
        flags=re.DOTALL
    )

    # Update default ID fallback
    first_pid = products[0]['id']
    prd_content = re.sub(
        r"const id = PRODUCTS\[params\.get\('id'\)\] \? params\.get\('id'\) : '[^']+';",
        f"const id = PRODUCTS[params.get('id')] ? params.get('id') : '{first_pid}';",
        prd_content
    )

    with open(PRODUK_FILE, 'w', encoding='utf-8') as f:
        f.write(prd_content)
    print("[OK] produk.html updated.")

    # -------------------------------------------------------------
    # D. UPDATE SITEMAP.XML
    # -------------------------------------------------------------
    print("Updating sitemap.xml...")
    sitemap_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url><loc>https://thetaleofher.com/</loc><priority>1.0</priority></url>
  <url><loc>https://thetaleofher.com/katalog.html</loc><priority>0.8</priority></url>
"""
    for p in products:
        sitemap_xml += f"  <url><loc>https://thetaleofher.com/produk.html?id={p['id']}&amp;c=0</loc><priority>0.6</priority></url>\n"
    sitemap_xml += "</urlset>\n"

    with open(SITEMAP_FILE, 'w', encoding='utf-8') as f:
        f.write(sitemap_xml)
    print("[OK] sitemap.xml updated.")

    print("\n=============================================================")
    print("    ALL PAGES SUCCESSFULLY UPDATED & SYNCHRONIZED!")
    print("=============================================================")

if __name__ == '__main__':
    build()
