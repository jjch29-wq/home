import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_cols_logic = """        self.ndt_cols = ('검사방법', '구간', '라인번호', 'Joint No.', '관경', '용접사', '구간정보', '결과', '규격', 
                         'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT')
        history = self.load_history()
        sections = set()
        lines = set()
        for date_str, data in history.items():
            for r in data.get('ndt_results', []):
                if r.get('구간'): sections.add(r['구간'].strip())
                if r.get('라인번호'): lines.add(r['라인번호'].strip())
        
        self.history_sections = [''] + sorted(list(sections))
        self.history_lines = [''] + sorted(list(lines))"""

new_cols_logic = """        self.ndt_cols = ('업체', '검사방법', '구간', '라인번호', 'Joint No.', '관경', '용접사', '구간정보', '결과', '규격', 
                         'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT')
        history = self.load_history()
        sections = set()
        lines = set()
        companies = set()
        for date_str, data in history.items():
            for r in data.get('ndt_results', []):
                if r.get('구간'): sections.add(r['구간'].strip())
                if r.get('라인번호'): lines.add(r['라인번호'].strip())
                if r.get('업체'): companies.add(r['업체'].strip())
        
        self.history_sections = [''] + sorted(list(sections))
        self.history_lines = [''] + sorted(list(lines))
        self.history_companies = [''] + sorted(list(companies))"""

code = code.replace(old_cols_logic, new_cols_logic)

old_width_logic = """                w = 8
                if c in ('검사방법', '결과', '규격', '관경'): w = 6
                elif c == '구간': w = 8
                elif c == '용접사': w = 15
                elif c == '라인번호': w = 25
                elif c == 'Joint No.': w = 12
                elif c == '구간정보': w = 20
                else: w = 5"""

new_width_logic = """                w = 8
                if c in ('검사방법', '결과', '규격', '관경'): w = 6
                elif c in ('구간', '업체'): w = 10
                elif c == '용접사': w = 15
                elif c == '라인번호': w = 25
                elif c == 'Joint No.': w = 12
                elif c == '구간정보': w = 20
                else: w = 5"""

code = code.replace(old_width_logic, new_width_logic)

old_entry_logic = """                elif c == '구간':
                    ent = ttk.Combobox(grid_frame, width=w, values=self.history_sections, justify='center')"""

new_entry_logic = """                elif c == '업체':
                    ent = ttk.Combobox(grid_frame, width=w, values=self.history_companies, justify='center')
                elif c == '구간':
                    ent = ttk.Combobox(grid_frame, width=w, values=self.history_sections, justify='center')"""

code = code.replace(old_entry_logic, new_entry_logic)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated daily_work_log_tab.py successfully")
