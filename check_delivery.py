"""Analyse delivery data across all orgs for Similarities/Differences/Gaps"""
import json

BASE = r'C:\Users\Administrator\Downloads\NHS Big List'
with open(f'{BASE}\\org_map.json', 'r', encoding='utf-8') as f:
    org_data = json.load(f)

# Collect all delivery areas per org
delivery_areas = {}  # area_name -> set of orgs that have it
dimension_coverage = {}  # dim_key -> set of orgs with data

for org_name, org in org_data.items():
    a = org.get('analysis')
    if not a:
        continue
    
    # Check dimensions
    dim_keys = ['GEOGRAPHY','10YP_ALIGNMENT','10YP_DELIVERABILITY','TRADE-OFFS','PREVENTION_READINESS','DIGITAL_READINESS','CARE_CLOSE_TO_HOME','WORKFORCE_READINESS','EQUITY_READINESS','INNOVATION_TRANSFORMATION','ADAPTABILITY',
                'SUMMARY','VISION','AMBITIONS','ENABLERS','DELIVERY_SHORT']
    for d in dim_keys:
        if a.get(d) and str(a[d]).strip():
            dimension_coverage.setdefault(d, set()).add(org_name)
    
    # Check DELIVERY table
    dtable = a.get('DELIVERY')
    if dtable and len(dtable) > 1:
        for row in dtable[1:]:  # skip header
            area = str(row[0]).strip() if row[0] else ''
            if area and area != 'None':
                delivery_areas.setdefault(area, set()).add(org_name)

print("=== DELIVERY AREA COVERAGE ===")
for area in sorted(delivery_areas, key=lambda a: -len(delivery_areas[a])):
    orgs = delivery_areas[area]
    print(f"{area}: {len(orgs)} orgs -> {', '.join(sorted(orgs)[:5])}{'...' if len(orgs)>5 else ''}")

print("\n=== DIMENSION COVERAGE ===")
for d in sorted(dimension_coverage, key=lambda x: -len(dimension_coverage[x])):
    orgs = dimension_coverage[d]
    print(f"{d}: {len(orgs)}/{len(org_data)} orgs")

# Find orgs missing delivery data
no_delivery = [n for n, o in org_data.items() 
               if not o.get('analysis') or not o['analysis'].get('DELIVERY')]
print(f"\n=== ORGS WITHOUT DELIVERY TABLE: {len(no_delivery)} ===")
for n in no_delivery:
    print(f"  {n}")
