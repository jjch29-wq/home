with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

code = code.replace(
    "ent = ttk.Combobox(grid_frame, width=w, values=['', 'RT', 'PAUT', 'UT', 'MT', 'PT', 'PMI', 'ETC'])",
    "ent = ttk.Combobox(grid_frame, width=w, values=['', 'RT', 'PAUT', 'UT', 'MT', 'PT', 'PMI', 'ETC'], justify='center')"
)
code = code.replace(
    "e = ttk.Entry(frame, width=3)",
    "e = ttk.Entry(frame, width=3, justify='center')"
)
code = code.replace(
    "ent = ttk.Combobox(grid_frame, width=w, values=[''] + list(SIZE_LENGTH.keys()))",
    "ent = ttk.Combobox(grid_frame, width=w, values=[''] + list(SIZE_LENGTH.keys()), justify='center')"
)
code = code.replace(
    "ent = ttk.Combobox(grid_frame, width=w, values=['', '합격', '불합격', '재촬영', '보류'])",
    "ent = ttk.Combobox(grid_frame, width=w, values=['', '합격', '불합격', '재촬영', '보류'], justify='center')"
)
code = code.replace(
    "ent = ttk.Combobox(grid_frame, width=w, values=[''])",
    "ent = ttk.Combobox(grid_frame, width=w, values=[''], justify='center')"
)
code = code.replace(
    "ent = ttk.Combobox(grid_frame, width=w, values=['', 'W-2023-A-10', 'W-2023-A-13', 'W-2023-A-25'])",
    "ent = ttk.Combobox(grid_frame, width=w, values=['', 'W-2023-A-10', 'W-2023-A-13', 'W-2023-A-25'], justify='center')"
)
code = code.replace(
    "ent = ttk.Entry(grid_frame, width=w)",
    "ent = ttk.Entry(grid_frame, width=w, justify='center')"
)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated successfully")
