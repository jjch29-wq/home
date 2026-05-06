import re

file_path = 'src/Material-Master-Manager-V13.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add missing bindings and defaults
additions = """
        # Restore focus transitions and defaults
        self.ent_daily_inspection_item.insert(0, "Piping")
        self.ent_daily_inspection_item.bind('<Return>', lambda e: self.ent_daily_test_amount.focus_set())
        self.ent_daily_test_amount.bind('<Return>', lambda e: self.cb_daily_unit.focus_set())
        self.cb_daily_unit.set('매')
        self.cb_daily_unit.bind('<Return>', lambda e: self.ent_daily_unit_price.focus_set())
        self.cb_daily_unit.bind('<<ComboboxSelected>>', lambda e: self.ent_daily_unit_price.focus_set())
        self.ent_daily_unit_price.bind('<Return>', lambda e: self.ent_daily_applied_code.focus_set())
        self.ent_daily_applied_code.insert(0, "KS")
        self.ent_daily_applied_code.bind('<Return>', lambda e: self.ent_daily_report_no.focus_set())
        self.ent_daily_report_no.bind('<Return>', lambda e: self.ent_daily_note.focus_set())
        self.ent_daily_note.bind('<Return>', lambda e: self.ent_daily_meal_cost.focus_set())
        self.ent_daily_meal_cost.insert(0, "0")
        self.ent_daily_meal_cost.bind('<Return>', lambda e: self.ent_daily_test_fee.focus_set())
        
        def on_method_select_focus(e):
            self.root.after(10, self.ent_daily_inspection_item.focus_set)
            return "break"
        self.cb_daily_test_method.bind('<<ComboboxSelected>>', on_method_select_focus, add='+')
        self.cb_daily_test_method.bind('<Return>', on_method_select_focus, add='+')
        
        def on_method_change_auto_unit_logic(e):
            method = self.cb_daily_test_method.get().strip()
            unit_map = {'RT': '매', 'UT': 'P,M,I/D', 'MT': 'P,M,I/D', 'PT': 'P,M,I/D', 'PAUT': 'M,I/D'}
            if method in unit_map: self.cb_daily_unit.set(unit_map[method])
            return "break"
        self.cb_daily_test_method.bind('<<ComboboxSelected>>', on_method_change_auto_unit_logic, add='+')
        
        self._bind_combobox_word_suggest(self.cb_daily_material, lambda: self._get_equipment_candidates(include_all=False))
"""

# Insert before '# Category definitions'
content = content.replace('# Category definitions', additions + "\n        # Category definitions")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Bindings and defaults restored.')
