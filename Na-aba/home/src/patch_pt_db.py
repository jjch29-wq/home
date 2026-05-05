import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# Find the chem/NDT block in export_daily_work_report
# It was recently patched to use index-based reading from UI. 
# We'll change it to DB-first.
start_anchor = "                # 화학약품 (MT/PT) - ndt_company_entries 위젯에서 순서 기반 읽기"
end_anchor   = "                        if used:\n                            data['materials'][m_key] = {'used': used}"

if start_anchor not in content:
    print("FAILED: start_anchor not found")
    exit(1)
if end_anchor not in content:
    print("FAILED: end_anchor not found")
    exit(1)

parts = content.split(start_anchor, 1)
prefix = parts[0]
rest = parts[1].split(end_anchor, 1)
suffix = rest[1]

# New block: Sum NDT_... columns from DB if site_records exists, else fallback to UI
new_chem_block = """                # [NEW] NDT 화학약품 (MT/PT) - DB site_records에서 합산하여 수집
                # 컬럼명 매핑: DB(NDT_...) -> materials key
                chem_db_map = [
                    ('MT WHITE',     'NDT_백색페인트'),
                    ('MT 7C-BLACK',  'NDT_흑색자분'),
                    ('PT Penetrant', 'NDT_침투제'),
                    ('PT Cleaner',   'NDT_세척제'),
                    ('PT Developer', 'NDT_현상제'),
                ]
                
                db_chem_found = False
                if not site_records.empty:
                    for m_key, db_col in chem_db_map:
                        if db_col in site_records.columns:
                            try:
                                val_sum = int(pd.to_numeric(site_records[db_col], errors='coerce').fillna(0).sum())
                                if val_sum > 0:
                                    data['materials'][m_key] = {'used': val_sum}
                                    db_chem_found = True
                            except: pass
                
                # DB에 데이터가 전혀 없으면 UI 위젯에서 읽기 (Fallback)
                if not db_chem_found and hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
                    entries_dict = self.ndt_company_entries[0]
                    mat_keys = [k for k in entries_dict.keys() if k != '_company']
                    fallback_order = [
                        ('MT WHITE', 2), ('MT 7C-BLACK', 1), 
                        ('PT Penetrant', 3), ('PT Cleaner', 4), ('PT Developer', 5)
                    ]
                    for m_key, idx in fallback_order:
                        if idx < len(mat_keys):
                            try:
                                val = entries_dict[mat_keys[idx]].get().strip()
                                used = int(val) if val else 0
                                if used > 0: data['materials'][m_key] = {'used': used}
                            except: pass"""

new_content = prefix + new_chem_block + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(new_content)
print("SUCCESS: PT/MT material collection from DB inserted")
