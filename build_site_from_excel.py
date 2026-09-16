import openpyxl
import json
import re
import os
import shutil

EXCEL_SRC = r'C:\Users\DTProject\Documents\Git\thetaleofher\Update\thetaleofher_content_template.xlsx'
EXCEL_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\thetaleofher_content_template.xlsx'
INDEX_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\index.html'
KATALOG_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\katalog.html'
PRODUK_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\produk.html'
SITEMAP_FILE = r'C:\Users\DTProject\Documents\Git\thetaleofher\sitemap.xml'

# Copy updated Excel from Update folder to root
if os.path.exists(EXCEL_SRC):
    shutil.copy2(EXCEL_SRC, EXCEL_FILE)

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

def parse_colors_str(colors_str):
    """
    Parses strings like:
    'White (#F2ECE1), Black (#241A1F)' -> [{'name': 'White', 'hex': '#F2ECE1'}, {'name': 'Black', 'hex': '#241A1F'}]
    'Light Blue' -> [{'name': 'Light Blue', 'hex': '#ADC7D9'}]
    """
    if not colors_str:
        return [{'name': 'Standard', 'hex': '#E5DCD3'}]
        
    parts = [p.strip() for p in colors_str.split(',') if p.strip()]
    results = []
    
    for p in parts:
        # Check for (#HEX)
        hex_match = re.search(r'#([0-9a-fA-F]{6}|[0-9a-fA-F]{3})', p)
        if hex_match:
            hex_val = '#' + hex_match.group(1).upper()
            cname = re.sub(r'\(#.*?\)', '', p).strip()
        else:
            cname = p.strip()
            # lookup in map
            hex_val = '#E5DCD3'
            for k, v in COLOR_HEX_MAP.items():
                if k in cname.lower():
                    hex_val = v
                    break
        results.append({'name': cname, 'hex': hex_val})
        
    return results if results else [{'name': 'Standard', 'hex': '#E5DCD3'}]

def build():
    print("Loading Excel template...")
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    ws_prod = wb['Products']
    products = []
    for r in ws_prod.iter_rows(min_row=2, values_only=True):
        if r[0]:
            pid = str(r[0]).strip().lower()
            name = str(r[1] or '').strip()
            cat = str(r[2] or '').strip().lower()
            price = int(r[3]) if r[3] else 0
            tag = str(r[4] or '').strip()
            color_raw = str(r[5] or '').strip()
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
            
            parsed_colors = parse_colors_str(color_raw)
            
            # Map gallery shots per color
            color_objs = []
            for c_idx, c_info in enumerate(parsed_colors):
                c_name_slug = c_info['name'].lower().replace(' ', '')
                # Filter shots matching color name
                matched_shots = [img for img in gallery if c_name_slug in img.lower()]
                if not matched_shots:
                    if len(parsed_colors) == 1:
                        matched_shots = gallery if gallery else [front_img]
                    else:
                        # Split gallery evenly if not matched
                        chunk_size = max(1, len(gallery) // len(parsed_colors))
                        matched_shots = gallery[c_idx*chunk_size : (c_idx+1)*chunk_size]
                        
                c_front = matched_shots[0] if matched_shots else front_img
                color_objs.append({
                    'name': c_info['name'],
                    'hex': c_info['hex'],
                    'img': c_front,
                    'shots': matched_shots if matched_shots else [c_front]
                })

            products.append({
                'id': pid,
                'name': name,
                'cat': cat,
                'price': price,
                'tag': tag,
                'colors': color_objs,
                'sizing': sizing,
                'material': material,
                'fit': fit,
                'desc': desc,
                'care': care,
                'status': status
            })

    print(f"Loaded {len(products)} products from Excel.")

    # -------------------------------------------------------------
    # A. UPDATE INDEX.HTML (Featured Grid)
    # -------------------------------------------------------------
    print("Updating index.html...")
    with open(INDEX_FILE, 'r', encoding='utf-8') as f:
        idx_content = f.read()

    featured = products[:4]
    feat_html_list = []
    delays = ['', 'd1', 'd2', 'd3']
    for i, p in enumerate(featured):
        d_class = delays[i] if i < len(delays) else ''
        first_color = p['colors'][0]
        feat_card = f'''      <a href="produk.html?id={p['id']}&c=0" class="prod reveal {d_class} group block">
        <div class="relative overflow-hidden aspect-[3/4] bg-[var(--paper-2)]">
          <img src="{first_color['img']}" alt="{p['name']} {first_color['name']}" class="w-full h-full object-cover object-top">
          <span class="quick absolute bottom-3 left-1/2 -translate-x-1/2 bg-[var(--paper)]/95 text-[var(--ink)] text-[10px] t2 uppercase px-5 py-2.5 whitespace-nowrap">View Details</span>
        </div>
        <div class="flex items-start justify-between mt-4">
          <div><h4 class="f text-lg leading-tight">{p['name']}</h4><p class="text-[12px] text-[var(--stone)] mt-0.5">{first_color['name']}</p></div>
          <p class="text-[13px] whitespace-nowrap">Rp {p['price']:,}</p>
        </div>
      </a>'''
        feat_html_list.append(feat_card)
    feat_grid_html = '\n\n'.join(feat_html_list)

    match = re.search(r'(<div class="grid grid-cols-2 lg:grid-cols-4 gap-x-5 gap-y-14">)(.*?)(</div>\s*</section>)', idx_content, re.DOTALL)
    if match:
        idx_content = idx_content[:match.start(2)] + '\n' + feat_grid_html + '\n    ' + idx_content[match.end(2):]

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

    kat_content = re.sub(
        r'<div class="flex items-center gap-2\.5 overflow-x-auto no-scrollbar">.*?</div>',
        lambda m: f'<div class="flex items-center gap-2.5 overflow-x-auto no-scrollbar">{cat_chips_html}</div>',
        kat_content
    )

    # Build products JS array for katalog.html
    # For each color variant, create a catalog entry
    js_prods = []
    for p in products:
        all_swatches = [c['hex'] for c in p['colors']]
        all_main_imgs = [c['img'] for c in p['colors']]
        
        for c_idx, c in enumerate(p['colors']):
            fu = c['img']
            bu = c['shots'][1] if len(c['shots']) > 1 else fu
            tag_prop = f"tag:'{p['tag']}', " if p['tag'] else ""
            sw_json = json.dumps(all_swatches)
            imgs_json = json.dumps(all_main_imgs)
            
            entry = f"    {{name:'{p['name']}', color:'{c['name']}', price:{p['price']}, cat:'{p['cat']}', {tag_prop}fu:'{fu}', bu:'{bu}', href:'produk.html?id={p['id']}&c={c_idx}', sw:{sw_json}, imgs:{imgs_json}}}"
            js_prods.append(entry)

    js_prods_str = 'const products = [\n' + ',\n'.join(js_prods) + '\n  ];'

    script_start_idx = kat_content.find('<script>')
    grid_idx = kat_content.find('const grid = document.getElementById(\'grid\');')
    if script_start_idx != -1 and grid_idx != -1:
        prefix = kat_content[:script_start_idx]
        suffix = kat_content[grid_idx:]
        new_script_head = f"""<script>
  const IMG = (id) => `https://images.unsplash.com/photo-${{id}}?fm=jpg&q=70&w=800&auto=format&fit=crop`;
  const rp = (n) => (n && n > 0) ? 'Rp ' + n.toLocaleString('id-ID') : 'Price on request';

  {js_prods_str}

  """
        kat_content = prefix + new_script_head + suffix

    with open(KATALOG_FILE, 'w', encoding='utf-8') as f:
        f.write(kat_content)
    print(f"[OK] katalog.html updated ({len(js_prods)} catalog cards generated for {len(products)} products).")

    # -------------------------------------------------------------
    # C. UPDATE PRODUK.HTML
    # -------------------------------------------------------------
    print("Updating produk.html...")
    with open(PRODUK_FILE, 'r', encoding='utf-8') as f:
        prd_content = f.read()

    prod_dict_entries = []
    related_entries = []

    for p in products:
        colors_json_parts = []
        for c_idx, c in enumerate(p['colors']):
            shots_json = json.dumps(c['shots'])
            colors_json_parts.append(f"{{name:'{c['name']}', hex:'{c['hex']}', img:'{c['img']}', shots:{shots_json}}}")
            related_entries.append(
                f"    {{id:'{p['id']}', c:{c_idx}, name:'{p['name']}', color:'{c['name']}', price:{p['price']}, img:'{c['img']}'}}"
            )
            
        colors_arr_str = "[ " + ", ".join(colors_json_parts) + " ]"
        mat_text = p['material']
        if p['care']:
            mat_text += f" {p['care']}"
            
        entry = f"""    {p['id']}: {{
      name: {json.dumps(p['name'])},
      price: {p['price']},
      line: 'The Tale of Her · The Classics',
      desc: {json.dumps(p['desc'])},
      fit: {json.dumps(p['fit'])},
      mat: {json.dumps(mat_text)},
      ship: 'Shipping: Orders placed before 04:00 p.m. will be processed on the same day.\\n\\nReturn: Exchanges accepted within 3 days with unboxing video via WhatsApp +62-818-0505-2929.',
      shopeeLink: '',
      colors: {colors_arr_str}
    }}"""
        prod_dict_entries.append(entry)

    products_obj_str = "const PRODUCTS = {\n" + ",\n".join(prod_dict_entries) + "\n  };"
    related_obj_str = "const RELATED = [\n" + ",\n".join(related_entries) + "\n  ];"

    p_script_idx = prd_content.find('<script>')
    p_params_idx = prd_content.find('const params = new URLSearchParams(location.search);')
    if p_script_idx != -1 and p_params_idx != -1:
        p_prefix = prd_content[:p_script_idx]
        p_suffix = prd_content[p_params_idx:]
        new_p_script_head = f"""<script>
  const rp = n => (n && n > 0) ? 'Rp ' + n.toLocaleString('id-ID') : 'Price on request';
  {products_obj_str}
  {related_obj_str}

  """
        prd_content = p_prefix + new_p_script_head + p_suffix

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
        for c_idx in range(len(p['colors'])):
            sitemap_xml += f"  <url><loc>https://thetaleofher.com/produk.html?id={p['id']}&amp;c={c_idx}</loc><priority>0.6</priority></url>\n"
    sitemap_xml += "</urlset>\n"

    with open(SITEMAP_FILE, 'w', encoding='utf-8') as f:
        f.write(sitemap_xml)
    print("[OK] sitemap.xml updated.")

    print("\n=============================================================")
    print(f"    SUCCESSFULLY COMPILED ALL {len(products)} PRODUCTS (16 TOTAL)!")
    print("=============================================================")

if __name__ == '__main__':
    build()
