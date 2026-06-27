import os

path = r'Material-Master-Manager-V13.py'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# We will replace the entire block between 11800 and 11950 with clean code
# This block should contain _save_ndt_product_map and open_ndt_product_map_dialog
new_block = '''    def _save_ndt_product_map(self, map_data):
        """NDT 약품 -> 실제 DB 품목명 매핑 저장"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
            else:
                cfg = {}
            cfg['ndt_product_map'] = map_data
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            return True
        except Exception:
            return False

    def open_ndt_product_map_dialog(self):
        """NDT 약품 -> 실제 DB 품목명 매핑 설정 다이얼로그"""
        dlg = tk.Toplevel(self.root)
        dlg.title("NDT 약품-품목 매핑 설정")
        dlg.geometry("500x600")
        dlg.transient(self.root)
        dlg.grab_set()

        main_frame = ttk.Frame(dlg, padding=20)
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="현장 입력 약품명", font=('Arial', 10, 'bold')).grid(row=0, column=0, pady=10)
        ttk.Label(main_frame, text="창고 재고 품목 (매핑)", font=('Arial', 10, 'bold')).grid(row=0, column=1, pady=10)

        ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        current_map = self._load_ndt_product_map()
        
        material_options = []
        for _, row in self.materials_df.iterrows():
            disp = self.get_material_display_name(row['MaterialID'])
            material_options.append(disp)
        material_options.sort()

        combos = {}
        for i, mat in enumerate(ndt_materials):
            ttk.Label(main_frame, text=mat).grid(row=i+1, column=0, padx=5, pady=5, sticky='w')
            cb = ttk.Combobox(main_frame, values=material_options, width=40)
            cb.grid(row=i+1, column=1, padx=5, pady=5, sticky='ew')
            current_id = current_map.get(mat, "")
            if current_id:
                cb.set(self.get_material_display_name(current_id))
            combos[mat] = cb

        def _save():
            new_map = {}
            for mat, cb in combos.items():
                disp = cb.get().strip()
                if disp:
                    for _, row in self.materials_df.iterrows():
                        if self.get_material_display_name(row['MaterialID']) == disp:
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

# We identify the start and end of the mess
# Start of mess is around 11801
# End of mess is just before def _get_merged_memo_and_note
block_start = -1
block_end = -1
for i, line in enumerate(lines[:13000]):
    if 'def _save_ndt_product_map' in line:
        block_start = i
    if 'def _get_merged_memo_and_note' in line:
        block_end = i
        break

if block_start != -1 and block_end != -1:
    print(f"Replacing block from line {block_start+1} to {block_end}")
    new_lines = lines[:block_start] + [new_block] + lines[block_end:]
    with open(path, 'w', encoding='utf-8', errors='ignore') as f:
        f.writelines(new_lines)
    print("SUCCESS: Block rebuilt")
else:
    print(f"FAILED: Could not identify block boundaries (start={block_start}, end={block_end})")
