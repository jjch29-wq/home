import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# 1. Fix the broken dialog block (from iterrows loop to end of button frame)
# We need to find where the mess started
mess_start_marker = '                        if self.get_material_display_name(row[\'MaterialID\'])'
mess_end_anchor = '    def _get_merged_memo_and_note(self):'

if mess_start_marker not in content:
    print("CRITICAL: Mess start not found")
    exit(1)

parts = content.split(mess_start_marker, 1)
prefix = parts[0]
remaining = parts[1].split(mess_end_anchor, 1)
suffix = remaining[1]

# Correct code for the end of _save and the buttons in open_ndt_product_map_dialog
correct_dialog_tail = ''' == disp:
                            new_map[mat] = row['MaterialID']
                            break
            if self._save_ndt_product_map(new_map):
                messagebox.showinfo("성공", "매핑 설정이 저장되었습니다.")
                dlg.destroy()
            else:
                messagebox.showerror("오류", "설정을 저장하지 못했습니다.")

        def _clear():
            for cb in combos.values():
                cb.set('')

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=len(ndt_materials)+1, column=0, columnspan=2, pady=20)

        ttk.Button(btn_frame, text="저장", command=_save, width=10).pack(side='left', padx=8)
        ttk.Button(btn_frame, text="전체 초기화", command=_clear, width=10).pack(side='left', padx=8)
        ttk.Button(btn_frame, text="닫기", command=dlg.destroy, width=10).pack(side='left', padx=8)

    '''

# Rebuild files without the mess
content = prefix + mess_start_marker + correct_dialog_tail + mess_end_anchor + suffix

with open(path, 'w', encoding='utf-8', errors='ignore') as f:
    f.write(content)

print("SUCCESS: Corrupted dialog block replaced with clean code.")
