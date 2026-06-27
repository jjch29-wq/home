import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# 1. Remove the bad block at 11962
bad_block = """            # [NEW] NDT 화학약품 (MT/PT) - DB site_records에서 합산하여 수집
            chem_db_map = [
                ('MT WHITE',     ['NDT_백색페인트', 'NDT_형광자분', 'NDT_백색페인트_MT']), # 기존/신규 컬럼 모두 대응
                ('MT 7C-BLACK',  ['NDT_흑색자분']),
                ('PT Penetrant', ['NDT_침투제']),
                ('PT Cleaner',   ['NDT_세척제']),
                ('PT Developer', ['NDT_현상제']),
            ]
            
            db_chem_found = False
            if not site_records.empty:
                for m_key, db_cols in chem_db_map:
                    val_sum = 0
                    for col in db_cols:
                        if col in site_records.columns:
                            try:
                                val_sum += int(pd.to_numeric(site_records[col], errors='coerce').fillna(0).sum())
                            except: pass
                    
                    if val_sum > 0:
                        data['materials'][m_key] = {'used': val_sum}
                        db_chem_found = True"""

if bad_block in content:
    content = content.replace(bad_block, "")
    print("SUCCESS: Bad block removed")

# 2. Correct the export function logic
correct_export_block = """            # [NEW] NDT 화학약품 (MT/PT) - DB site_records에서 합산하여 수집
            chem_db_map = [
                ('MT WHITE',     ['NDT_백색페인트', 'NDT_형광자분', 'NDT_백색페인트_MT']), 
                ('MT 7C-BLACK',  ['NDT_흑색자분']),
                ('PT Penetrant', ['NDT_침투제']),
                ('PT Cleaner',   ['NDT_세척제']),
                ('PT Developer', ['NDT_현상제']),
            ]
            
            db_chem_found = False
            if not site_records.empty:
                for m_key, db_cols in chem_db_map:
                    val_sum = 0
                    for col in db_cols:
                        if col in site_records.columns:
                            try:
                                val_sum += int(pd.to_numeric(site_records[col], errors='coerce').fillna(0).sum())
                            except: pass
                    
                    if val_sum > 0:
                        data['materials'][m_key] = {'used': val_sum}
                        db_chem_found = True
            
            # DB에 데이터가 없으면 UI 위젯에서 읽기 (Fallback)
            if not db_chem_found and hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
                entries_dict = self.ndt_company_entries[0]
                mat_keys = [k for k in entries_dict.keys() if k != '_company']
                fallback_order = [
                    ('MT WHITE', 0), ('MT 7C-BLACK', 1), 
                    ('PT Penetrant', 3), ('PT Cleaner', 4), ('PT Developer', 5)
                ]
                for m_key, idx in fallback_order:
                    if idx < len(mat_keys):
                        try:
                            val = entries_dict[mat_keys[idx]].get().strip()
                            used = int(val) if val else 0
                            if used > 0: data['materials'][m_key] = {'used': used}
                        except: pass"""

# Find the old UI-only block to replace
old_ui_anchor = "            # 화학약품 (MT, PT) - UI에서 읽기"
if old_ui_anchor in content:
    # Find the next "data['rtk'] = {}" and replace everything between
    parts = content.split(old_ui_anchor, 1)
    subparts = parts[1].split("            data['rtk'] = {}", 1)
    content = parts[0] + correct_export_block + "\n\n            data['rtk'] = {}" + subparts[1]
    print("SUCCESS: Export function logic updated")

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(content)
