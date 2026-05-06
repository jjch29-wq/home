import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# Find the Name/Spec lookup block inside the mat_groups loop
start_anchor = "                    if hasattr(self, 'materials_df') and not self.materials_df.empty:\n                        try:\n                            match = self.materials_df["
end_anchor   = "                        mat_groups[mat_id_upper] = {\n                            'qty': 0, 'name': mat_name_val, 'spec': mat_spec_val, 'original_id': mat_id\n                        }"

if start_anchor not in content:
    print("FAILED: start_anchor not found")
    exit(1)
if end_anchor not in content:
    print("FAILED: end_anchor not found")
    exit(1)

parts = content.split(start_anchor, 1)
prefix = parts[0]
rest   = parts[1].split(end_anchor, 1)
suffix = rest[1]

# New logic: parse MaterialID directly by splitting on '-'
# e.g. "Carestream AA400-3⅓*12"  ->  name="Carestream AA400", spec="3⅓*12"
new_parse_block = """                    # MaterialID를 '-' 기준으로 품목명 / 규격 분리
                    # 예: "Carestream AA400-3⅓*12"  ->  name="Carestream AA400", spec="3⅓*12"
                    if '-' in mat_id:
                        dash_idx   = mat_id.index('-')
                        mat_name_val = mat_id[:dash_idx].strip()
                        mat_spec_val = mat_id[dash_idx+1:].strip()
                    else:
                        mat_name_val = mat_id
                        mat_spec_val = ''

                    mat_groups[mat_id_upper] = {
                        'qty': 0, 'name': mat_name_val, 'spec': mat_spec_val, 'original_id': mat_id
                    }"""

new_content = prefix + new_parse_block + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(new_content)
print("SUCCESS: MaterialID parsing updated (split on '-')")
