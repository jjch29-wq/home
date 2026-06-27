import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

start_anchor = "                        # MaterialID를 '-' 기준으로 분리"
end_anchor = "                    mat_groups[mat_id_upper]['qty'] += qty"

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

new_block = """                        # DB의 MaterialID(숫자 혹은 코드)로 materials_df에서 품목명, 규격 조회
                        mat_name_val = str(mat_id)
                        mat_spec_val = ''
                        disp_name = ''  # For _mat_key logic
                        if hasattr(self, 'materials_df') and not self.materials_df.empty:
                            try:
                                match = self.materials_df[
                                    self.materials_df['MaterialID'].astype(str).str.strip().str.replace(r'\\.0$', '', regex=True) == str(mat_id).strip().replace('.0', '')
                                ]
                                if not match.empty:
                                    mat_name_val = str(match.iloc[0].get('품목명', mat_name_val)).strip()
                                    mat_spec_val = str(match.iloc[0].get('규격', '')).strip()
                                    if mat_spec_val in ('nan', 'None'): mat_spec_val = ''
                            except: pass

                        disp_name = f"{mat_name_val}-{mat_spec_val}".upper() if mat_spec_val else mat_name_val.upper()

                        mat_groups[mat_id_upper] = {
                            'qty': 0, 'name': mat_name_val, 'spec': mat_spec_val, 'original_id': mat_id, 'disp': disp_name
                        }
"""

new_content = prefix + new_block + end_anchor + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(new_content)
print("SUCCESS: lookup logic added")
