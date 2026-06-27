import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

start_anchor = "                        # Look up Name and Spec from materials_df"
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

new_block = """                        # MaterialID를 '-' 기준으로 분리 (예: Carestream AA400-3⅓*12 -> Carestream AA400, 3⅓*12)
                        mat_name_val = mat_id
                        mat_spec_val = ''
                        if '-' in mat_id:
                            dash_idx = mat_id.index('-')
                            mat_name_val = mat_id[:dash_idx].strip()
                            mat_spec_val = mat_id[dash_idx+1:].strip()

                        mat_groups[mat_id_upper] = {
                            'qty': 0, 'name': mat_name_val, 'spec': mat_spec_val, 'original_id': mat_id
                        }
"""

new_content = prefix + new_block + end_anchor + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(new_content)
print("SUCCESS: split logic added")
