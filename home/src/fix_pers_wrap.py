import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_code = """        set_cell('O9', '구분 (관리/안전)', font=self.font_small, fill=self.fill_header)
        set_cell('P9', '검사원', font=self.font_bold, fill=self.fill_header)
        
        personnel = data.get('personnel_data', {})
        set_cell('O10', '인원')
        set_cell('P10', personnel.get('검사원_인원', ''))
        set_cell('O11', '현장대리인')
        set_cell('P11', personnel.get('검사원_현장대리인', ''))
        set_cell('O12', '누계')
        set_cell('P12', personnel.get('검사원_누계', ''))"""

new_code = """        set_cell('O9', '구분(관리/안전)', font=self.font_small, fill=self.fill_header, align=self.align_nowrap)
        set_cell('P9', '검사원', font=self.font_bold, fill=self.fill_header, align=self.align_nowrap)
        
        personnel = data.get('personnel_data', {})
        set_cell('O10', '인원', align=self.align_nowrap)
        set_cell('P10', personnel.get('검사원_인원', ''), align=self.align_nowrap)
        set_cell('O11', '현장대리인', align=self.align_nowrap)
        set_cell('P11', personnel.get('검사원_현장대리인', ''), align=self.align_nowrap)
        set_cell('O12', '누계', align=self.align_nowrap)
        set_cell('P12', personnel.get('검사원_누계', ''), align=self.align_nowrap)"""

code = code.replace(old_code, new_code)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated personnel alignment successfully")
