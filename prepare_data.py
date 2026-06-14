import openpyxl, json, sys
sys.stdout.reconfigure(encoding='utf-8')

# Load extracted data
with open(r'C:\Users\Administrator\Downloads\NHS Big List\analysis_data.json', 'r', encoding='utf-8') as f:
    analysis_data = json.load(f)
with open(r'C:\Users\Administrator\Downloads\NHS Big List\contacts_data.json', 'r', encoding='utf-8') as f:
    contacts_data = json.load(f)

# Merge into org_map
canon = {'Mersey Care NHS foundation Trust': 'Mersey Care NHS Foundation Trust',
    'Northern Care Alliance NHS Group ': 'Northern Care Alliance NHS Group',
    'East Cheshire': 'East Cheshire NHS Trust', 'East Lancs': 'East Lancashire Hospitals NHS Trust',
    'Wirral University Teaching Hospital NHS Foundation Trust ': 'Wirral University Teaching Hospital NHS Foundation Trust',
    'Bridgewater Community Healthcare NHS Foundation Trust': 'Bridgewater Community Healthcare NHS Foundation Trust',
    'Stockport NHS Foundation Trust  & Tameside & Glossop Integrated Care NHS Foundation Trust': 'Stockport NHS Foundation Trust'}
rnorm = {'LANCS': 'LSC'}

org_map = {}
for c in contacts_data:
    org = canon.get(c['org'], c['org'])
    region = rnorm.get(c['region'], c['region'])
    if org not in org_map:
        org_map[org] = {'org': org, 'region': region, 'type': c['type'], 'contacts': []}
    else:
        if region and not org_map[org]['region']: org_map[org]['region'] = region
        if not org_map[org]['type'] and c['type']: org_map[org]['type'] = c['type']
    dup = any(ex['name'] == c['name'] and ex['email'] == c['email'] and ex['job'] == c['job'] for ex in org_map[org]['contacts'])
    if not dup and c['name']: org_map[org]['contacts'].append(c)

# Analysis mappings
amap = {
    "Alder Hey Children's NHSFT": "Alder Hey Children's NHS Foundation Trust",
    'Blackpool Teaching NHSFT': 'Blackpool Teaching Hospital NHS Foundation Trust',
    'Bolton NHS Foundation Trust': 'Bolton NHS Foundation Trust',
    'Cheshire and Wirral Partnership NHS Foundation Trust': 'Cheshire and Wirral Partnership NHS Foundation Trust',
    'Clatterbridge Cancer Centre NHS Foundation Trust': 'Clatterbridge Cancer Centre NHS Foundation Trust',
    'Countess of Chester Hospital NHSFT': 'Countess of Chester Hospital NHS Foundation Trust',
    'East Cheshire NHS Trust': 'East Cheshire NHS Trust',
    'East Lancashire Hospitals NHST': 'East Lancashire Hospitals NHS Trust',
    'Greater Manchester Mental Health NHS Foundation Trust': 'Greater Manchester Mental Health NHS Foundation Trust',
    'Lancashire Teaching Hospitals NHSFT': 'Lancashire Teaching Hospitals NHS Foundation Trust',
    'Lancashire and South Cumbria NHS Foundation Trust': 'Lancashire and South Cumbria NHS Foundation Trust',
    'Manchester University NHS Foundation Trust': 'Manchester University NHS Foundation Trust',
    'Mersey Care NHS Foundation Trust': 'Mersey Care NHS Foundation Trust',
    'Mersey and West Lancashire Teaching Hospitals NHS Trust': 'Mersey and West Lancashire Teaching Hospitals NHS Trust',
    'Mid Cheshire Hospital NHS Foundation Trust': 'Mid Cheshire Hospitals NHS Foundation Trust',
    'North Cheshire and Mersey NHSFT - Warrington and Halton Teaching Hospitals NHSFT (WHH) and Bridgewater Community Healthcare NHSFT (BCH) will become a single organisation from April 2026': 'Warrington and Halton Hospitals NHS Foundation Trust',
    'North West Ambulance Service NHS Trust': 'North West Ambulance Service NHS Trust',
    'Northern Care Alliance NHS Foundation Trust': 'Northern Care Alliance NHS Group',
    'Pennine Care NHS Foundation Trust': 'Pennine Care NHS Foundation Trust',
    'Stockport NHS Foundation Trust (SFT) & Tameside & Glossop Integrated Care NHS Foundation Trust (TGICFT)': 'Stockport NHS Foundation Trust',
    'Tameside & Glossop Integrated Care NHS Foundation Trust (TGICFT)': 'Tameside & Glossop Integrated Care NHS Foundation Trust',
    'The Christie NHS Foundation Trust': 'The Christie NHS Foundation Trust',
    'The Walton Centre NHSFT': 'The Walton Centre NHS Foundation Trust',
    'University Hospitals of Morecambe Bay NHSFT': 'University Hospitals of Morecambe Bay NHS Foundation Trust',
    'Wirral Community Health and Care NHS FT\nand\nWirral University Teaching Hospital NHS FT': 'Wirral Community Health & Care NHS Foundation Trust',
    'Wrightington, Wigan and Leigh Teaching Hospitals NHS Foundation Trust': 'Wrightington, Wigan and Leigh NHS Foundation Trust',
}

for a_name, a_data in sorted(analysis_data.items()):
    target = amap.get(a_name, None)
    if target is None:
        continue
    if target in org_map:
        org_map[target]['analysis'] = a_data['analysis']
        if not org_map[target]['region'] and a_data['region']:
            org_map[target]['region'] = a_data['region'].replace('C&amp;M', 'C&M')
    else:
        r = a_data.get('region', '').replace('C&amp;M', 'C&M')
        org_map[a_name] = {'org': a_name, 'region': r, 'type': '', 'contacts': [], 'analysis': a_data['analysis']}

lg = analysis_data.get("University Hospitals of Liverpool Group (UHLG) currently comprised of Liverpool University Hospital (LUHFT), Liverpool Women\u2019s (LWH) and Liverpool Heart and Chest (LHCH)", None)
if lg:
    for lo in ['Liverpool University Hospitals NHS Foundation Trust', "Liverpool Women's Hospital NHS Foundation Trust", 'Liverpool Heart and Chest Hospital NHS Foundation Trust']:
        if lo in org_map:
            org_map[lo]['analysis'] = lg['analysis']
            if not org_map[lo]['region']: org_map[lo]['region'] = 'C&M'

# Compute risks
risk_items = []
kw = ['risk', 'challenge', 'concern', 'issue', 'uncertainty', 'threat', 'vulnerability', 'gap', 'shortfall', 'deficit', 'pressure']
for on, od in sorted(org_map.items()):
    if not od.get('analysis'): continue
    a = od['analysis']
    for sname in ['TRADE-OFFS', 'WORKFORCE_READINESS', 'ADAPTABILITY', 'SUMMARY', 'OTHER']:
        text = ''
        if sname == 'OTHER' and isinstance(a.get(sname), list):
            text = ' '.join(a[sname])
        elif isinstance(a.get(sname), str):
            text = a[sname]
        if text:
            for sent in text.replace('\n', ' ').split('. '):
                if any(w in sent.lower() for w in kw):
                    risk_items.append({'org': on, 'section': sname, 'text': sent.strip()})

# Stats
systems = {}
type_counts = {}
for k, v in org_map.items():
    s = v.get('region', 'Unknown')
    systems[s] = systems.get(s, 0) + 1
    t = v.get('type', 'Unknown')
    type_counts[t] = type_counts.get(t, 0) + 1

cps = {}
for k, v in org_map.items():
    s = v.get('region', 'Unknown')
    cps[s] = cps.get(s, 0) + len(v.get('contacts', []))

dims = ['GEOGRAPHY','10YP_ALIGNMENT','10YP_DELIVERABILITY','TRADE-OFFS','PREVENTION_READINESS','DIGITAL_READINESS','CARE_CLOSE_TO_HOME','WORKFORCE_READINESS','EQUITY_READINESS','INNOVATION_TRANSFORMATION','ADAPTABILITY']
dc = {}
for d in dims:
    dc[d] = sum(1 for v in org_map.values() if v.get('analysis', {}).get(d))

stats = {
    'systems': systems, 'type_counts': type_counts,
    'total_orgs': len(org_map), 'total_contacts': sum(len(v['contacts']) for v in org_map.values()),
    'orgs_with_analysis': sum(1 for v in org_map.values() if v.get('analysis')),
    'contacts_per_system': cps, 'dim_coverage': dc
}

# Write JSON data files
with open(r'C:\Users\Administrator\Downloads\NHS Big List\org_map.json', 'w', encoding='utf-8') as f:
    json.dump(org_map, f, ensure_ascii=False)
with open(r'C:\Users\Administrator\Downloads\NHS Big List\risks.json', 'w', encoding='utf-8') as f:
    json.dump(risk_items, f, ensure_ascii=False)
with open(r'C:\Users\Administrator\Downloads\NHS Big List\stats.json', 'w', encoding='utf-8') as f:
    json.dump(stats, f, ensure_ascii=False)

print('Data files written.')
print(f'Orgs: {len(org_map)}')
print(f'Contacts: {sum(len(v["contacts"]) for v in org_map.values())}')
print(f'Risks: {len(risk_items)}')
