import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# Find the anchor to insert RTK aggregation block BEFORE the file dialog
anchor = '            # 3. ???\xec\xa0\x80\xec\x9e\xa5 \xea\xb2\xbd\xeb\xa1\x9c \xec\xa0\x84\xec\xa0\x95'

# Try multiple potential Korean strings for the save dialog comment
anchors_to_try = [
    '            # 3. 저장경로 전정',
    '            # 3. ???',
    'default_filename = f"',
]

found_anchor = None
for a in anchors_to_try:
    if a in content:
        found_anchor = a
        break

# Just find 'default_filename' which is unique enough
if 'default_filename = f"' in content:
    found_anchor = '            default_filename = f"'

if not found_anchor:
    print("Could not find insertion anchor. Trying 'save_path = filedialog'...")
    if 'save_path = filedialog.asksaveasfilename(' in content:
        found_anchor = '            save_path = filedialog.asksaveasfilename('

if not found_anchor:
    print("FAILED: No anchor found")
    exit(1)

rtk_block = '''            # [NEW] RTK 불량 정보 DB에서 집계 (현장 탭 기록기준)
            data['rtk'] = {}
            rtk_cats = {
                '센터미스': 'center_miss', '농도': 'density', '마킹미스': 'marking_miss',
                '필름마크': 'film_mark', '취급부주의': 'handling', '고객불만': 'customer_complaint', '기타': 'etc'
            }
            rtk_total = 0
            for kor_key in rtk_cats.keys():
                db_val = 0
                if not site_records.empty:
                    col = f"RTK_{kor_key}"
                    if col in site_records.columns:
                        db_val = int(pd.to_numeric(site_records[col], errors='coerce').fillna(0).sum())
                data['rtk'][kor_key] = db_val
                rtk_total += db_val
            data['rtk']['총계'] = rtk_total

            '''

new_content = content.replace(found_anchor, rtk_block + found_anchor, 1)
if new_content == content:
    print("FAILED: Replace had no effect")
    exit(1)

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(new_content)
print("SUCCESS: RTK DB aggregation block inserted")
