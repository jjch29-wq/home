import pandas as pd
df = pd.read_excel('c:/Users/-/PMI/home/src/site_apps/central/data/Material_Inventory.xlsx', sheet_name='DailyUsage')
v_map = {}
for _, row in df.iterrows():
    raw_v_no = str(row.get('차량번호','')).strip()
    date_val = str(row.get('Date', str(row.get('일자', '')))).strip()[:10]
    if raw_v_no and raw_v_no.lower() not in ['nan','none','']:
        v_list = [v.strip() for v in raw_v_no.replace('||',',').split(',') if v.strip()]
        for v in v_list:
            v_map.setdefault(v, set()).add(date_val)

print("Vehicle dates mapping:")
for v, dates in v_map.items():
    print(f"  {v}: {sorted(list(dates))} ({len(dates)} days)")
total = sum(len(d) for d in v_map.values())
print("Total vehicle days:", total)
