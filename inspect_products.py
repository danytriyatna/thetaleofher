import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook(r'C:\Users\DTProject\Documents\Git\thetaleofher\Update\thetaleofher_content_template.xlsx', data_only=True)
ws = wb['Products']
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[0]:
        pid = str(r[0]).strip()
        name = str(r[1] or '').strip()
        colors = str(r[5] or '').strip()
        front = str(r[11] or '').strip()
        back = str(r[12] or '').strip()
        gallery = [g.strip() for g in str(r[13] or '').split(',') if g.strip()]
        print(f"[{pid}] {name} | Colors: {colors} | Front: {front} | Back: {back} | Gallery: {len(gallery)} imgs")
