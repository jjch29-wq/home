import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# Find the broken chem_map block anchor
# The block starts with "chem_map = ["
start_anchor = '                chem_map = ['
end_anchor   = '                        data[\'materials\'][m_key] = {\'used\': used}'

if start_anchor not in content:
    print("FAILED: start_anchor not found")
    exit(1)
if end_anchor not in content:
    print("FAILED: end_anchor not found")
    exit(1)

# split
parts = content.split(start_anchor, 1)
prefix = parts[0]
rest = parts[1].split(end_anchor, 1)
suffix = rest[1]

# New chem block: iterate widget entries by position (0=형광자분, 1=흑색자분, 2=백색페인트, 3=침투제, 4=세척제, 5=현상제, 6=형광침투제)
# Map index -> materials key
new_chem_block = """                # 화학약품 (MT/PT) - ndt_company_entries 위젯에서 순서 기반 읽기
                chem_order = [
                    ('MT WHITE',     0),   # 형광자분 (index 0)
                    ('MT 7C-BLACK',  1),   # 흑색자분 (index 1)
                    ('PT Penetrant', 3),   # 침투제   (index 3)
                    ('PT Cleaner',   4),   # 세척제   (index 4)
                    ('PT Developer', 5),   # 현상제   (index 5)
                ]
                if hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
                    entries_dict = self.ndt_company_entries[0]
                    # Get list of mat widgets (excluding _company key)
                    mat_keys = [k for k in entries_dict.keys() if k != '_company']
                    for m_key, idx in chem_order:
                        used = 0
                        if idx < len(mat_keys):
                            try:
                                val = entries_dict[mat_keys[idx]].get().strip()
                                used = int(val) if val else 0
                            except: used = 0
                        if used:
                            data['materials'][m_key] = {'used': used}"""

# rebuild
new_content = prefix + new_chem_block + '\n' + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(new_content)
print("SUCCESS: Chem map fixed to use index-based widget reading")
