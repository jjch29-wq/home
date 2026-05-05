import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# Find the material section anchor: "data['selected_material']"
anchor = "data['selected_material'] = self.cb_daily_material.get().strip()"
if anchor not in content:
    print("FAILED: anchor not found")
    exit(1)

# Find the end of the UI-based material section (up to RTK block)
end_anchor = "data['rtk'] = {}"
if end_anchor not in content:
    print("FAILED: end_anchor not found")
    exit(1)

# Split
parts = content.split(anchor, 1)
prefix = parts[0]
rest = parts[1].split(end_anchor, 1)
middle = rest[0]   # the old UI-based material code (will be replaced)
suffix = rest[1]

# New material block: DB-first (group by MaterialID), fallback to UI
new_material_block = """data['selected_material'] = self.cb_daily_material.get().strip()

            # [NEW] RT 품목별 수량 - DB site_records의 MaterialID로 그룹화하여 배정
            # 매핑: T200 -> data['materials']['RT T200'], AA400 -> data['materials']['RT AA400'], 기타 -> 'RT Other'
            if not site_records.empty and 'MaterialID' in site_records.columns:
                # MaterialID 기준으로 Usage 합산
                mat_groups = {}
                for _, row in site_records.iterrows():
                    mat_id = str(row.get('MaterialID', '')).strip()
                    qty = 0
                    try: qty = float(row.get('Usage', 0) or 0)
                    except: pass
                    if mat_id and mat_id != 'nan':
                        mat_id_upper = mat_id.upper()
                        if mat_id_upper not in mat_groups:
                            mat_groups[mat_id_upper] = 0
                        mat_groups[mat_id_upper] += qty

                # 매핑 테이블: 품목명 키워드 -> materials 키
                def _mat_key(mat_id_upper):
                    if 'T200' in mat_id_upper:      return 'RT T200'
                    if 'AA400' in mat_id_upper:     return 'RT AA400'
                    if 'AA' in mat_id_upper:        return 'RT AA400'
                    # 필름 종류이면 RT T200 기본
                    return 'RT Other'

                for mat_id_upper, qty_sum in mat_groups.items():
                    mat_key = _mat_key(mat_id_upper)
                    if mat_key not in data['materials']:
                        data['materials'][mat_key] = {'used': 0}
                    data['materials'][mat_key]['used'] += int(qty_sum)

            else:
                # DB 데이터 없으면 UI에서 읽기 (fallback)
                if hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
                    mats = self.ndt_company_entries[0]
                    for m_key, mat_name in [('RT T200', 'T200'), ('RT AA400', 'AA400')]:
                        if m_key in mats:
                            val = mats[m_key].get().strip()
                            try: used = int(val) if val else 0
                            except: used = 0
                            data['materials'][m_key] = {'used': used}

            # 화학약품 (MT, PT) - UI에서 읽기
            if hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
                mats = self.ndt_company_entries[0]
                chem_map = [
                    ('MT WHITE',      '형광자분'),  ('MT 7C-BLACK',   '흑색자분'),
                    ('PT Penetrant',  '침투제'),    ('PT Cleaner',    '세척제'),
                    ('PT Developer',  '현상제'),
                ]
                for m_key, chem_key in chem_map:
                    if chem_key in mats:
                        try:
                            val = mats[chem_key].get().strip()
                            used = int(val) if val else 0
                        except: used = 0
                        data['materials'][m_key] = {'used': used}

            """

# Rebuild
new_content = prefix + new_material_block + end_anchor + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(new_content)
print("SUCCESS: RT material grouping by MaterialID inserted")
