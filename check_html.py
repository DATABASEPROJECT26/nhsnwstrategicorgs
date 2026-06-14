import re, os, json

f = r'C:\Users\Administrator\Downloads\NHS Big List\database.html'
with open(f, 'r', encoding='utf-8') as fh:
    html = fh.read()

checks = [
    ('Title', 'NHS Northwest Organisation Database'),
    ('Chart.js CDN', 'cdn.jsdelivr.net/npm/chart.js'),
    ('Leaflet CSS', 'leaflet@1.9/dist/leaflet.css'),
    ('Leaflet JS', 'leaflet@1.9/dist/leaflet.js'),
    ('Add button', 'class="add-btn"'),
    ('Add menu', 'showAddMenu'),
    ('Add org', 'showAddOrg'),
    ('Add contact', 'showAddContactTo'),
    ('Add field', 'showAddField'),
    ('Add region', 'showAddRegion'),
    ('Add type', 'showAddType'),
    ('Submit org', 'submitOrg'),
    ('Submit contact', 'submitContactTo'),
    ('Submit field', 'submitField'),
    ('Delete contact', 'deleteContact'),
    ('Export JSON', 'exportData'),
    ('Map container', 'map-container'),
    ('Map render', 'renderMap'),
    ('Get Directions', 'google.com/maps/dir'),
    ('Directions button', 'dir-btn'),
    ('Leaflet marker', 'L.marker'),
    ('Leaflet tileLayer', 'L.tileLayer'),
    ('Pencil icon', '&#9998;'),
    ('Edit function', 'editField'),
    ('Edit analysis', 'editAnalysisField'),
    ('Modal overlay', 'modal-overlay'),
    ('Modal content', 'modalContent'),
    ('Toast', 'toast'),
    ('getFilteredOrgs', 'getFilteredOrgs'),
    ('Dashboard respects filters', 'renderDashboard(filtered)'),
    ('Risks respects filters', 'renderRisks(filtered)'),
]

all_ok = True
for name, pattern in checks:
    ok = pattern in html
    if not ok:
        print(f'MISSING: {name} ({pattern})')
        all_ok = False
if all_ok:
    print('All 30 features present in database.html')
else:
    print('SOME CHECKS FAILED')

print(f'Total: {len(html)} bytes, {html.count(chr(10))} lines')

# Verify data blob is valid JSON
import re
m = re.search(r'<script id="data-blob" type="application/json">(.*?)</script>', html, re.DOTALL)
if m:
    try:
        data = json.loads(m.group(1))
        print(f'Data JSON valid: {len(data)} orgs')
        latlng_count = sum(1 for k in data if data[k].get('latlng'))
        print(f'  with latlng: {latlng_count}')
    except Exception as e:
        print(f'Data JSON INVALID: {e}')
