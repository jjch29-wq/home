import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update _build_right_pane to extract history
start_marker = "        self.ndt_cols = ('검사방법', '구간', '라인번호', 'Joint No.', '관경', '용접사', '구간정보', '결과', '규격', \n                         'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT')"
history_extract = """
        history = self.load_history()
        sections = set()
        lines = set()
        for date_str, data in history.items():
            for r in data.get('ndt_results', []):
                if r.get('구간'): sections.add(r['구간'].strip())
                if r.get('라인번호'): lines.add(r['라인번호'].strip())
        
        self.history_sections = [''] + sorted(list(sections))
        self.history_lines = [''] + sorted(list(lines))
"""
code = code.replace(start_marker, start_marker + history_extract)

# 2. Update widget creation
old_else = """                else:
                    ent = ttk.Entry(grid_frame, width=w, justify='center')
                    ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    row_entries[c] = ent"""

new_else = """                elif c == '구간':
                    ent = ttk.Combobox(grid_frame, width=w, values=self.history_sections, justify='center')
                    ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    row_entries[c] = ent
                elif c == '라인번호':
                    ent = ttk.Combobox(grid_frame, width=w, values=self.history_lines, justify='center')
                    ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    row_entries[c] = ent
                else:
                    ent = ttk.Entry(grid_frame, width=w, justify='center')
                    ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    row_entries[c] = ent"""

code = code.replace(old_else, new_else)

# 3. Update save_current_history
old_save = """            data['ndt_results'].append(row_dict)
            
        history[current_date] = data"""

new_save = """            data['ndt_results'].append(row_dict)
            
            # Update dynamic combobox lists
            if row_dict.get('구간') and row_dict['구간'] not in self.history_sections:
                self.history_sections.append(row_dict['구간'])
            if row_dict.get('라인번호') and row_dict['라인번호'] not in self.history_lines:
                self.history_lines.append(row_dict['라인번호'])
                
        # Sort and refresh comboboxes
        if '' in self.history_sections: self.history_sections.remove('')
        if '' in self.history_lines: self.history_lines.remove('')
        self.history_sections = [''] + sorted(self.history_sections)
        self.history_lines = [''] + sorted(self.history_lines)
        
        for row_entries in self.ndt_grid_entries:
            if hasattr(row_entries['구간'], 'configure'):
                row_entries['구간']['values'] = self.history_sections
            if hasattr(row_entries['라인번호'], 'configure'):
                row_entries['라인번호']['values'] = self.history_lines
            
        history[current_date] = data"""

code = code.replace(old_save, new_save)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Added history comboboxes successfully")
