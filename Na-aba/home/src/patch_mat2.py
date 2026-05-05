import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# ========== Find and replace the RT material grouping block in export function ==========
# Anchor: the loop that builds mat_groups
start_anchor = "            # [NEW] RT 품목별 수량 - DB site_records의 MaterialID로 그룹화하여 배정"
end_anchor   = "            else:\n                # DB 데이터 없으면 UI에서 읽기 (fallback)\n                if hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:\n                    mats = self.ndt_company_entries[0]\n                    for m_key, mat_name in [('RT T200', 'T200'), ('RT AA400', 'AA400')]:\n                        if m_key in mats:\n                            val = mats[m_key].get().strip()\n                            try: used = int(val) if val else 0\n                            except: used = 0\n                            data['materials'][m_key] = {'used': used}"

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

new_rt_block = """            # [NEW] RT 품목별 수량 - DB site_records의 MaterialID로 그룹화 + 품목명/규격 조회
            # D열: 품목명(Name), F열: 규격(Spec), M열: 사용수량(Usage)
            if not site_records.empty and 'MaterialID' in site_records.columns:
                # MaterialID 기준으로 Usage 합산 및 Name/Spec 조회
                mat_groups = {}  # key: MaterialID_upper -> {'qty': float, 'name': str, 'spec': str, 'original_id': str}
                for _, row in site_records.iterrows():
                    mat_id = str(row.get('MaterialID', '')).strip()
                    if not mat_id or mat_id == 'nan':
                        continue
                    qty = 0
                    try: qty = float(row.get('Usage', 0) or 0)
                    except: pass

                    mat_id_upper = mat_id.upper()
                    if mat_id_upper not in mat_groups:
                        # Look up Name and Spec from materials_df
                        mat_name_val = mat_id   # fallback: use ID as name
                        mat_spec_val = ''
                        if hasattr(self, 'materials_df') and not self.materials_df.empty:
                            try:
                                match = self.materials_df[
                                    self.materials_df['MaterialID'].astype(str).str.strip() == mat_id
                                ]
                                if match.empty:
                                    # Try name-based lookup
                                    match = self.materials_df[
                                        self.materials_df['Name'].astype(str).str.strip().str.upper() == mat_id_upper
                                    ]
                                if not match.empty:
                                    mat_name_val = str(match.iloc[0].get('Name', mat_id)).strip()
                                    mat_spec_val = str(match.iloc[0].get('Spec', '')).strip()
                                    if mat_spec_val in ('nan', 'None', 'NULL', ''): mat_spec_val = ''
                            except: pass

                        mat_groups[mat_id_upper] = {
                            'qty': 0, 'name': mat_name_val, 'spec': mat_spec_val, 'original_id': mat_id
                        }
                    mat_groups[mat_id_upper]['qty'] += qty

                # Map to data['materials'] keys (T200->row43, AA400->row44, others->row45)
                def _mat_key(mid_upper):
                    if 'T200'  in mid_upper: return 'RT T200'
                    if 'AA400' in mid_upper: return 'RT AA400'
                    if 'AA'    in mid_upper: return 'RT AA400'
                    return 'RT Other'

                for mat_id_upper, grp in mat_groups.items():
                    mat_key = _mat_key(mat_id_upper)
                    if mat_key not in data['materials']:
                        data['materials'][mat_key] = {'used': 0, 'name': grp['name'], 'spec': grp['spec']}
                    else:
                        data['materials'][mat_key]['used'] = data['materials'][mat_key].get('used', 0)
                        if not data['materials'][mat_key].get('name'):
                            data['materials'][mat_key]['name'] = grp['name']
                        if not data['materials'][mat_key].get('spec'):
                            data['materials'][mat_key]['spec'] = grp['spec']
                    data['materials'][mat_key]['used'] = data['materials'][mat_key].get('used', 0) + int(grp['qty'])

            else:
                # DB 데이터 없으면 UI에서 읽기 (fallback)
                if hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
                    mats = self.ndt_company_entries[0]
                    for m_key, mat_name in [('RT T200', 'T200'), ('RT AA400', 'AA400')]:
                        if m_key in mats:
                            val = mats[m_key].get().strip()
                            try: used = int(val) if val else 0
                            except: used = 0
                            data['materials'][m_key] = {'used': used}"""

new_content = prefix + new_rt_block + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(new_content)
print("SUCCESS: RT material block updated with Name/Spec lookup")
