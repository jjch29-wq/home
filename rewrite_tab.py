import os
import re

file_path = r"c:\Users\jjch2\Desktop\PMI\home\src\daily_work_log_tab.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update qty_rows
old_qty_rows = """        self.qty_rows = [
            ('PAUT', '300A이상'), ('PAUT', '300A이상-야간'), ('PAUT', '250A'), ('PAUT', '200A'), ('PAUT', '200A-야간'), ('PAUT', '소계'),
            ('RT', '150A~100A'), ('RT', '150A~100A-야간'), ('RT', '80A이하'), ('RT', '80A이하-야간'), ('RT', '소계'),
            ('MT', '전체(주간)'), ('MT', '전체(야간)'),
            ('PT', '전체(주간)'), ('PT', '전체(야간)')
        ]"""
new_qty_rows = """        self.qty_rows = [
            ('RT', 'B필름: 3⅓"x17"'),
            ('RT', 'A필름: 3⅓"x12"'),
            ('RT', 'A/2필름: 3⅓"x6"'),
            ('RT', '소계'),
            ('UT', '초음파탐상'),
            ('PT', '침투탐상'),
            ('MT', '자분탐상')
        ]"""
content = content.replace(old_qty_rows, new_qty_rows)

# 2. Update default_qty
old_default_qty = """        self.default_qty = {
            ('PAUT', '300A이상'): '121', ('PAUT', '300A이상-야간'): '584',
            ('PAUT', '250A'): '4', ('PAUT', '200A'): '4',
            ('PAUT', '200A-야간'): '2', ('PAUT', '소계'): '715',
            ('RT', '150A~100A'): '293', ('RT', '150A~100A-야간'): '43',
            ('RT', '80A이하'): '105', ('RT', '80A이하-야간'): '49',
            ('RT', '소계'): '490', ('MT', '전체(주간)'): '26',
            ('MT', '전체(야간)'): '0', ('PT', '전체(주간)'): '26',
            ('PT', '전체(야간)'): '0',
        }"""
new_default_qty = """        self.default_qty = {
            ('RT', 'B필름: 3⅓"x17"'): '20368',
            ('RT', 'A필름: 3⅓"x12"'): '2464',
            ('RT', 'A/2필름: 3⅓"x6"'): '1704',
            ('RT', '소계'): '24536',
            ('UT', '초음파탐상'): '319.02',
            ('PT', '침투탐상'): '338.63',
            ('MT', '자분탐상'): '0'
        }"""
content = content.replace(old_default_qty, new_default_qty)

# 3. Update equip_rows
old_equip_rows = "        self.equip_rows = ['PAUT장비', 'PAUT프로브', 'PAUT스캐너', 'RT장비', 'MT장비']"
new_equip_rows = "        self.equip_rows = ['RT장비(선원)', 'UT장비', 'MT/PT장비']"
content = content.replace(old_equip_rows, new_equip_rows)

# 4. Update ndt_cols (Line 243)
old_ndt_cols = "        self.ndt_cols = ('업체', '검사방법', '구간', '라인번호', 'Joint No.', '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분',\n                         'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT')"
new_ndt_cols = "        self.ndt_cols = ('업체', '검사방법', '구간', '라인번호', 'Joint No.', '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분',\n                         'RT_OR', 'RT_RE', 'UT', 'MT', 'PT')"
if old_ndt_cols not in content:
    # Try one line
    old_ndt_cols_2 = "self.ndt_cols = ('업체', '검사방법', '구간', '라인번호', 'Joint No.', '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분', 'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT')"
    content = content.replace(old_ndt_cols_2, "self.ndt_cols = ('업체', '검사방법', '구간', '라인번호', 'Joint No.', '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분', 'RT_OR', 'RT_RE', 'UT', 'MT', 'PT')")
content = content.replace(old_ndt_cols, new_ndt_cols)

# 5. Update auto_calculate_and_save logic
# Find auto_calculate_and_save and replace its content
start_marker = "    def auto_calculate_and_save(self):"
end_marker = "        # Update 공정률 in UI"
# Actually, the end is where it updates "공정률" or "Update Totals based on previous date"
# Let's use regex to replace the body of auto_calculate_and_save

new_auto_calc = """    def auto_calculate_and_save(self):
        current_date = self.date_entry.get()
        history = self.load_history()
        
        # 1. Aggregate NDT Results -> 금일작업
        today_qty = {comp_key: 0.0 for comp_key in self.qty_entries.keys()}
        
        for row_entries in self.ndt_grid_entries:
            if not hasattr(row_entries['검사방법'], 'get'): continue
            method = row_entries['검사방법'].get().upper().strip()
            if not method: continue
            
            spec_str = row_entries['규격'].get().strip() if hasattr(row_entries['규격'], 'get') else ""
            
            if method == 'RT':
                val = float(row_entries['RT_OR'].get() or 0) + float(row_entries['RT_RE'].get() or 0)
                if val > 0:
                    if '17' in spec_str or 'B' in spec_str.upper():
                        spec_key = 'B필름: 3⅓"x17"'
                    elif '12' in spec_str:
                        spec_key = 'A필름: 3⅓"x12"'
                    elif '6' in spec_str:
                        spec_key = 'A/2필름: 3⅓"x6"'
                    else:
                        spec_key = 'B필름: 3⅓"x17"' # Default
                    comp = f"RT_{spec_key}"
                    if comp in today_qty: today_qty[comp] += val
            
            elif method == 'UT':
                val = float(row_entries['UT'].get() or 0)
                if val > 0:
                    comp = "UT_초음파탐상"
                    if comp in today_qty: today_qty[comp] += val
                        
            elif method == 'MT':
                val = float(row_entries['MT'].get() or 0)
                if val > 0:
                    comp = "MT_자분탐상"
                    if comp in today_qty: today_qty[comp] += val
                    
            elif method == 'PT':
                val = float(row_entries['PT'].get() or 0)
                if val > 0:
                    comp = "PT_침투탐상"
                    if comp in today_qty: today_qty[comp] += val
                    
        # Subtotals
        today_qty['RT_소계'] = sum([v for k, v in today_qty.items() if k.startswith('RT_') and k != 'RT_소계'])
                    
        # Update 금일작업 in UI
        def format_val(ckey, v):
            if ckey.startswith(('UT', 'PT')):
                return f"{v:.4f}"
            return f"{v:.1f}" if v % 1 else f"{int(v)}"

        for comp_key, val in today_qty.items():
            if '소계' not in comp_key or comp_key == 'RT_소계':
                ent = self.qty_entries[comp_key]['금일작업']
                ent.delete(0, tk.END)
                if val != 0:
                    ent.insert(0, format_val(comp_key, val))
                
        # 2. Update Totals based on previous date"""

# Use string matching
parts = content.split("    def auto_calculate_and_save(self):")
if len(parts) == 2:
    part2 = parts[1]
    part2_split = part2.split("        # 2. Update Totals based on previous date")
    if len(part2_split) >= 2:
        content = parts[0] + new_auto_calc + "        # 2. Update Totals based on previous date" + "        # 2. Update Totals based on previous date".join(part2_split[1:])
    else:
        print("Could not find Update Totals based on previous date")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print("Updated successfully")
