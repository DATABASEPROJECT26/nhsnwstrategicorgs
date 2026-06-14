import openpyxl, json, sys
sys.stdout.reconfigure(encoding='utf-8')

# Read one analysis sheet and compare with HTML
wb = openpyxl.load_workbook(r'C:\Users\Administrator\Downloads\NHS Big List\ANALYSIS OF NARRATIVE PLANS.xlsx', data_only=True, read_only=True)

# Pick a sheet with significant content: NW_LSC_NWAS (the big one)
ws = wb['NW_LSC_NWAS']
rows = list(ws.iter_rows(values_only=True))

# Verify row positions by printing headers for rows 11-14
for offset in range(10, 15):
    if offset < len(rows):
        print(f'Row {offset+1} header: {rows[offset][1]}')

# GEOGRAPHY is at row 12 (1-indexed) = rows[11] (0-indexed)
geo_source = str(rows[11][2]) if len(rows) >= 12 and rows[11][2] else ''
print(f'\nGEOGRAPHY length: {len(geo_source)} chars')
print(f'First 100: {geo_source[:100]}')
print()

# 10YP ALIGNMENT at row 13 (1-indexed) = rows[12] (0-indexed)
align_source = str(rows[12][2]) if len(rows) >= 13 and rows[12][2] else ''
print(f'10YP_ALIGNMENT length: {len(align_source)} chars')
print(f'First 100: {align_source[:100]}')
print(f'Source 10YP_ALIGNMENT length: {len(align_source)} chars')
print(f'First 100: {align_source[:100]}')
print(f'Last 100: {align_source[-100:]}')
print()

# Check delivery section size
delivery_start = None
for i, r in enumerate(rows):
    if r[1] and str(r[1]).strip().upper() == 'DELIVERY':
        delivery_start = i
        break
if delivery_start:
    delivery_text = ''
    for r in rows[delivery_start:]:
        for v in r:
            if v and str(v).strip():
                delivery_text += str(v).strip() + ' '
    print(f'Delivery text length: {len(delivery_text)} chars')

print()
print('Now checking HTML...')
with open(r'C:\Users\Administrator\Downloads\NHS Big List\database.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Extract JSON
start = html.find('var ORGS = ') + len('var ORGS = ')
depth = 0
end = start
for i in range(start, len(html)):
    if html[i] == '{': depth += 1
    elif html[i] == '}': depth -= 1
    if depth == 0: end = i + 1; break
html_orgs = json.loads(html[start:end])

# Find NWAS
nwas = None
for k, v in html_orgs.items():
    if 'Ambulance' in k:
        nwas = v
        break

if nwas and nwas.get('analysis'):
    a = nwas['analysis']
    if a.get('GEOGRAPHY'):
        print(f'HTML GEOGRAPHY length: {len(a["GEOGRAPHY"])} chars')
        print(f'First 100: {a["GEOGRAPHY"][:100]}')
    if a.get('10YP_ALIGNMENT'):
        print(f'HTML 10YP_ALIGNMENT length: {len(a["10YP_ALIGNMENT"])} chars')
        print(f'First 100: {a["10YP_ALIGNMENT"][:100]}')
    if a.get('DELIVERY'):
        delivery_json = json.dumps(a['DELIVERY'], ensure_ascii=False)
        print(f'HTML DELIVERY length: {len(delivery_json)} chars')
