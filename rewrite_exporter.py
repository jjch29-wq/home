import os
import re

file_path = r"c:\Users\jjch2\Desktop\PMI\home\src\daily_work_log_exporter.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update qty_rows (Lines ~113)
old_qty_rows = """        qty_rows = [
            ('PAUT', '300A이상'), ('PAUT', '300A이상-야간'), ('PAUT', '250A'), ('PAUT', '200A'), ('PAUT', '200A-야간'), ('PAUT', '소계'),
            ('RT', '150A~100A'), ('RT', '150A~100A-야간'), ('RT', '80A이하'), ('RT', '80A이하-야간'), ('RT', '소계'),
            ('MT', '전체(주간)'), ('MT', '전체(야간)'),
            ('PT', '전체(주간)'), ('PT', '전체(야간)')
        ]"""
new_qty_rows = """        qty_rows = [
            ('RT', 'B필름: 3⅓"x17"'),
            ('RT', 'A필름: 3⅓"x12"'),
            ('RT', 'A/2필름: 3⅓"x6"'),
            ('RT', '소계'),
            ('UT', '초음파탐상'),
            ('PT', '침투탐상')
        ]"""
content = content.replace(old_qty_rows, new_qty_rows)

# 2. Update merged cells (Lines ~162)
old_merge_qty = """        ws.merge_cells('A9:A14') # PAUT
        ws.merge_cells('A15:A19') # RT
        ws.merge_cells('A20:A21') # MT
        ws.merge_cells('A22:A23') # PT"""
new_merge_qty = """        ws.merge_cells('A9:A12') # RT"""
content = content.replace(old_merge_qty, new_merge_qty)

# 3. Update equip_rows (Line 176)
old_equip_rows = "        equip_rows = ['PAUT장비', 'PAUT프로브', 'PAUT스캐너', 'RT장비', 'MT장비']"
new_equip_rows = "        equip_rows = ['RT장비(선원)', 'UT장비', 'MT/PT장비']"
content = content.replace(old_equip_rows, new_equip_rows)

# 4. Update PAUT -> UT in columns
content = content.replace("'PAUT(m)'", "'UT(m)'")
content = content.replace("['PAUT', 'PT', 'MT']", "['UT', 'PT', 'MT']")
content = content.replace("res.get('PAUT', '')", "res.get('UT', '')")
content = content.replace("'PAUT_300A_D'", "'UT'")

# 5. Project Name
content = content.replace("용 역 명 : 2026년 중앙지사 열수송관 비파괴검사용역", "용 역 명 : 가산~가평 천연가스 공급시설 건설공사 비파괴검사기술용역")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print("Updated exporter successfully")
