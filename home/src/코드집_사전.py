import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "codebook_db.json")

class CodebookApp:
    def __init__(self, root):
        self.root = root
        self.root.title("코드집 사전 (Codebook Manager)")
        self.root.geometry("850x650")
        
        style = ttk.Style()
        style.theme_use('clam')
        
        self.data = []
        self.load_data()
        
        self.create_widgets()
        self.refresh_list()
        
    def load_data(self):
        if os.path.exists(DB_FILE):
            try:
                with open(DB_FILE, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
            except Exception as e:
                messagebox.showerror("로딩 오류", f"DB 파일을 불러오지 못했습니다:\n{e}")
                self.data = []
        else:
            self.data = [
                {"category": "예시(NDT 규격)", "find": "ASME Sec.V", "replace": "ISO 10863"},
                {"category": "예시(프로젝트명)", "find": "기존프로젝트명", "replace": "가산~가평 천연가스 공급시설"},
                {"category": "예시(현장용어)", "find": "차폐복", "replace": "차폐체"}
            ]

    def save_data(self):
        try:
            with open(DB_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("저장 오류", f"DB 파일을 저장하지 못했습니다:\n{e}")

    def create_widgets(self):
        # 상단 검색 및 필터 프레임
        search_frame = ttk.LabelFrame(self.root, text="🔍 검색 및 필터")
        search_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(search_frame, text="카테고리:").pack(side="left", padx=5, pady=5)
        self.combo_filter_cat = ttk.Combobox(search_frame, state="readonly", width=15)
        self.combo_filter_cat.pack(side="left", padx=5, pady=5)
        self.combo_filter_cat.bind("<<ComboboxSelected>>", lambda e: self.refresh_list())
        
        ttk.Label(search_frame, text="검색어:").pack(side="left", padx=(15, 5), pady=5)
        self.entry_search = ttk.Entry(search_frame, width=30)
        self.entry_search.pack(side="left", padx=5, pady=5)
        self.entry_search.bind("<KeyRelease>", lambda e: self.refresh_list())
        
        ttk.Button(search_frame, text="초기화", command=self.reset_filter).pack(side="left", padx=10, pady=5)

        # 메인 리스트 및 상세 보기 팬(PanedWindow)
        paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 왼쪽 리스트 프레임
        list_frame = ttk.Frame(paned)
        paned.add(list_frame, weight=5)
        
        # 오른쪽 상세 보기 프레임
        detail_frame = ttk.LabelFrame(paned, text="📖 상세 내용 보기")
        paned.add(detail_frame, weight=3)
        
        self.detail_text = tk.Text(detail_frame, wrap="word", font=("맑은 고딕", 11))
        
        detail_scrollbar = ttk.Scrollbar(detail_frame, orient="vertical", command=self.detail_text.yview)
        detail_scrollbar.pack(side="right", fill="y")
        self.detail_text.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.detail_text.config(yscrollcommand=detail_scrollbar.set)
        self.detail_text.config(state="disabled")
        
        columns = ("category", "find", "replace")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        self.tree.heading("category", text="분류 (카테고리)")
        self.tree.heading("find", text="찾을 내용 (기존 코드/문구)")
        self.tree.heading("replace", text="바꿀 내용 (새로운 코드/문구)")
        
        self.tree.column("category", width=150, anchor="center")
        self.tree.column("find", width=300)
        self.tree.column("replace", width=300)
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.config(yscrollcommand=scrollbar.set)
        
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        # 버튼 프레임 (가운데)
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(btn_frame, text="📋 찾을 내용 복사", command=lambda: self.copy_to_clipboard("find")).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="📋 바꿀 내용 복사", command=lambda: self.copy_to_clipboard("replace")).pack(side="left", padx=5)
        
        ttk.Button(btn_frame, text="📤 현재 목록을 통합 허브용(JSON)으로 내보내기", command=self.export_preset).pack(side="right", padx=5)

        # 하단 입력 및 수정 프레임
        input_frame = ttk.LabelFrame(self.root, text="✍️ 코드 추가 및 수정")
        input_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(input_frame, text="카테고리:").grid(row=0, column=0, padx=5, pady=10, sticky="e")
        self.combo_cat = ttk.Combobox(input_frame, width=15)
        self.combo_cat.grid(row=0, column=1, padx=5, pady=10, sticky="w")
        
        ttk.Label(input_frame, text="찾을 내용:").grid(row=0, column=2, padx=5, pady=10, sticky="e")
        self.entry_find = ttk.Entry(input_frame, width=25)
        self.entry_find.grid(row=0, column=3, padx=5, pady=10, sticky="w")
        
        ttk.Label(input_frame, text="바꿀 내용:").grid(row=0, column=4, padx=5, pady=10, sticky="e")
        self.entry_replace = ttk.Entry(input_frame, width=25)
        self.entry_replace.grid(row=0, column=5, padx=5, pady=10, sticky="w")
        
        ttk.Label(input_frame, text="코드 내용\n(상세 설명):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.text_details_input = tk.Text(input_frame, width=75, height=4, font=("맑은 고딕", 10))
        self.text_details_input.grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky="w")
        
        action_frame = ttk.Frame(input_frame)
        action_frame.grid(row=2, column=0, columnspan=6, pady=10)
        
        ttk.Button(action_frame, text="✨ 추가", command=self.add_code, width=15).pack(side="left", padx=10)
        ttk.Button(action_frame, text="💾 수정", command=self.update_code, width=15).pack(side="left", padx=10)
        ttk.Button(action_frame, text="🗑️ 삭제", command=self.delete_code, width=15).pack(side="left", padx=10)

    def update_categories(self):
        cats = sorted(list(set([d.get("category", "") for d in self.data if d.get("category", "")])))
        self.combo_cat['values'] = cats
        self.combo_filter_cat['values'] = ["전체"] + cats
        
        # 필터창에 선택된 값이 더 이상 존재하지 않으면 '전체'로 초기화
        if self.combo_filter_cat.get() not in ["전체"] + cats:
            self.combo_filter_cat.set("전체")

    def refresh_list(self):
        self.update_categories()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        filter_cat = self.combo_filter_cat.get()
        search_kw = self.entry_search.get().lower()
        
        for idx, row in enumerate(self.data):
            cat = row.get("category", "")
            f_txt = row.get("find", "")
            r_txt = row.get("replace", "")
            
            if filter_cat and filter_cat != "전체" and cat != filter_cat:
                continue
                
            if search_kw:
                if search_kw not in cat.lower() and search_kw not in f_txt.lower() and search_kw not in r_txt.lower():
                    continue
                    
            self.tree.insert("", "end", iid=str(idx), values=(cat, f_txt, r_txt))

    def reset_filter(self):
        self.combo_filter_cat.set("전체")
        self.entry_search.delete(0, tk.END)
        self.refresh_list()

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if selected:
            idx = int(selected[0])
            row = self.data[idx]
            
            # 하단 입력창 업데이트
            self.combo_cat.set(row.get("category", ""))
            self.entry_find.delete(0, tk.END)
            self.entry_find.insert(0, row.get("find", ""))
            self.entry_replace.delete(0, tk.END)
            self.entry_replace.insert(0, row.get("replace", ""))
            
            self.text_details_input.delete("1.0", tk.END)
            self.text_details_input.insert("1.0", row.get("details", ""))
            
            # 우측 상세 보기 업데이트
            cat_text = row.get("category", "")
            f_text = row.get("find", "")
            r_text = row.get("replace", "")
            d_text = row.get("details", "")
            
            detail_content = f"■ 분류 (카테고리)\n{cat_text}\n\n"
            detail_content += f"■ 찾을 내용 (설명/기존문구)\n{f_text}\n\n"
            detail_content += f"■ 바꿀 내용 (적용할 규격/코드)\n{r_text}\n\n"
            detail_content += f"■ 실제 코드 내용 (상세 설명)\n{d_text if d_text else '(등록된 상세 내용이 없습니다.)'}"
            
            self.detail_text.config(state="normal")
            self.detail_text.delete("1.0", tk.END)
            self.detail_text.insert("1.0", detail_content)
            self.detail_text.config(state="disabled")

    def add_code(self):
        cat = self.combo_cat.get().strip()
        f_txt = self.entry_find.get().strip()
        r_txt = self.entry_replace.get().strip()
        d_txt = self.text_details_input.get("1.0", tk.END).strip()
        
        if not f_txt:
            messagebox.showwarning("입력 오류", "찾을 내용을 입력해주세요.")
            return
            
        self.data.append({
            "category": cat,
            "find": f_txt,
            "replace": r_txt,
            "details": d_txt
        })
        self.save_data()
        self.refresh_list()
        
        # 입력창 및 미리보기 초기화
        self.combo_cat.set("")
        self.entry_find.delete(0, tk.END)
        self.entry_replace.delete(0, tk.END)
        self.text_details_input.delete("1.0", tk.END)
        
        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.config(state="disabled")
        
        self.entry_find.focus()

    def update_code(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("선택 오류", "수정할 코드를 위 목록에서 선택해주세요.")
            return
            
        idx = int(selected[0])
        cat = self.combo_cat.get().strip()
        f_txt = self.entry_find.get().strip()
        r_txt = self.entry_replace.get().strip()
        d_txt = self.text_details_input.get("1.0", tk.END).strip()
        
        if not f_txt:
            messagebox.showwarning("입력 오류", "찾을 내용을 입력해주세요.")
            return
            
        self.data[idx] = {
            "category": cat,
            "find": f_txt,
            "replace": r_txt,
            "details": d_txt
        }
        self.save_data()
        self.refresh_list()
        
    def delete_code(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("선택 오류", "삭제할 코드를 위 목록에서 선택해주세요.")
            return
            
        if messagebox.askyesno("삭제 확인", "선택한 코드를 정말 삭제하시겠습니까?"):
            # 다중 선택을 지원할 경우 뒤에서부터 삭제해야 인덱스가 안 꼬임
            idxs = sorted([int(s) for s in selected], reverse=True)
            for idx in idxs:
                del self.data[idx]
                
            self.save_data()
            self.refresh_list()
            
            self.combo_cat.set("")
            self.entry_find.delete(0, tk.END)
            self.entry_replace.delete(0, tk.END)
            self.text_details_input.delete("1.0", tk.END)
            
            self.detail_text.config(state="normal")
            self.detail_text.delete("1.0", tk.END)
            self.detail_text.config(state="disabled")

    def copy_to_clipboard(self, field="find"):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("안내", "복사할 항목을 먼저 선택해주세요.")
            return
            
        idx = int(selected[0])
        text = self.data[idx].get(field, "")
        
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        
        messagebox.showinfo("복사 완료", f"클립보드에 복사되었습니다:\n\n{text}")

    def export_preset(self):
        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("경고", "내보낼 항목이 없습니다.")
            return
            
        preset_data = []
        for item in items:
            val = self.tree.item(item, 'values')
            preset_data.append({"find": val[1], "replace": val[2]})
            
        filepath = filedialog.asksaveasfilename(
            title="절차서 수정 헬퍼용 단어 목록으로 내보내기",
            defaultextension=".json",
            filetypes=[("JSON 파일", "*.json")],
            initialfile="내보낸_코드세트.json"
        )
        
        if filepath:
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(preset_data, f, ensure_ascii=False, indent=4)
                messagebox.showinfo("내보내기 완료", f"성공적으로 내보냈습니다!\n절차서 수정 헬퍼의 [목록 불러오기]에서 사용하세요.\n\n저장 경로: {filepath}")
            except Exception as e:
                messagebox.showerror("오류", f"저장 중 오류 발생:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CodebookApp(root)
    root.mainloop()
