import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

try:
    from tkinterweb import HtmlFrame
except (ImportError, OSError):
    HtmlFrame = None

try:
    import mammoth
except ImportError:
    mammoth = None

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
        
        # 탭(Notebook) 구성
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)
        
        self.tab_code = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_code, text="코드 관리")
        
        self.tab_viewer = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_viewer, text="절차서 뷰어 (웹/HTML)")
        
        # PDF 뷰어 관련 상태 변수
        self.pdf_doc = None
        self.current_page = 0
        self.pdf_image_id = None
        self.pdf_photo = None
        
        self.create_widgets()
        self.create_viewer_widgets()
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
        search_frame = ttk.LabelFrame(self.tab_code, text="🔍 검색 및 필터")
        
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
        paned = ttk.PanedWindow(self.tab_code, orient=tk.HORIZONTAL)
        
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
        btn_frame = ttk.Frame(self.tab_code)
        
        ttk.Button(btn_frame, text="📋 찾을 내용 복사", command=lambda: self.copy_to_clipboard("find")).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="📋 바꿀 내용 복사", command=lambda: self.copy_to_clipboard("replace")).pack(side="left", padx=5)
        
        ttk.Button(btn_frame, text="📤 현재 목록을 통합 허브용(JSON)으로 내보내기", command=self.export_preset).pack(side="right", padx=5)

        # 하단 입력 및 수정 프레임
        input_frame = ttk.LabelFrame(self.tab_code, text="✍️ 코드 추가 및 수정")
        
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

        # === 레이아웃 배치 (화면 잘림 방지) ===
        search_frame.pack(side="top", fill="x", padx=10, pady=10)
        input_frame.pack(side="bottom", fill="x", padx=10, pady=10)
        btn_frame.pack(side="bottom", fill="x", padx=10, pady=5)
        paned.pack(side="top", fill="both", expand=True, padx=10, pady=5)

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

    def create_viewer_widgets(self):
        ctrl_frame = ttk.Frame(self.tab_viewer)
        ctrl_frame.pack(side="top", fill="x", padx=10, pady=5)
        
        ttk.Button(ctrl_frame, text="📂 워드 문서 열기 (.docx)", command=self.load_document).pack(side="left", padx=5)
        
        # Reference 추출 버튼
        ttk.Button(ctrl_frame, text="✨ 2.0 Reference 추출", command=self.extract_references).pack(side="left", padx=10)
        
        # 원본 열기 (수정용) 버튼
        self.btn_edit_doc = ttk.Button(ctrl_frame, text="📝 원본 워드로 열어서 직접 수정하기", command=self.open_current_document, state="disabled")
        self.btn_edit_doc.pack(side="left", padx=10)
        
        if HtmlFrame is None or mammoth is None:
            ttk.Label(ctrl_frame, text="⚠️ tkinterweb 또는 mammoth 라이브러리가 필요합니다.", foreground="red").pack(side="right", padx=10)
            
        # 뷰어 내 검색용 프레임
        search_frame = ttk.Frame(self.tab_viewer)
        search_frame.pack(side="top", fill="x", padx=10, pady=0)
        
        ttk.Label(search_frame, text="🔍 뷰어 내 텍스트 검색:").pack(side="left", padx=5)
        self.entry_viewer_search = ttk.Entry(search_frame, width=30)
        self.entry_viewer_search.pack(side="left", padx=5)
        self.entry_viewer_search.bind("<Return>", lambda e: self.search_in_viewer())
        
        ttk.Button(search_frame, text="검색 (다음)", command=self.search_in_viewer).pack(side="left", padx=2)
        ttk.Button(search_frame, text="초기화", command=self.reset_viewer_search).pack(side="left", padx=2)
        
        self.lbl_viewer_search_result = ttk.Label(search_frame, text="")
        self.lbl_viewer_search_result.pack(side="left", padx=10)
        
        self.current_search_index = 0
        self.last_search_query = ""
        
        viewer_frame = ttk.Frame(self.tab_viewer)
        viewer_frame.pack(side="top", fill="both", expand=True, padx=10, pady=5)
        
        if HtmlFrame:
            try:
                self.html_viewer = HtmlFrame(viewer_frame, messages_enabled=False)
                self.html_viewer.pack(fill="both", expand=True)
            except OSError:
                self.html_viewer = None
        else:
            self.html_viewer = None
            
        self.current_doc_text = ""

    def load_document(self):
        if self.html_viewer is None or mammoth is None:
            messagebox.showerror("오류", "문서를 렌더링하기 위한 필수 라이브러리가 없습니다.")
            return
            
        filepath = filedialog.askopenfilename(
            title="절차서 문서 열기 (Word)",
            filetypes=[("워드 파일", "*.docx"), ("모든 파일", "*.*")]
        )
        
        if not filepath:
            return
            
        if filepath.lower().endswith('.docx'):
            try:
                # HTML 변환 및 텍스트 데이터 보관
                with open(filepath, "rb") as docx_file:
                    result = mammoth.convert_to_html(docx_file)
                    html = result.value
                    
                    docx_file.seek(0)
                    text_result = mammoth.extract_raw_text(docx_file)
                    self.current_doc_text = text_result.value
                    
                # 깔끔한 렌더링을 위한 스타일 추가
                styled_html = f"""
                <html>
                <head>
                <style>
                    body {{ font-family: 'Malgun Gothic', sans-serif; line-height: 1.6; padding: 20px; }}
                    table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
                    th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                    th {{ background-color: #f2f2f2; }}
                    img {{ max-width: 100%; height: auto; }}
                </style>
                </head>
                <body>
                {html}
                </body>
                </html>
                """
                self.html_viewer.load_html(styled_html)
                self.notebook.update()
                
                # 원본 열기 버튼 활성화 및 경로 저장
                self.current_filepath = filepath
                self.btn_edit_doc.config(state="normal")
                
            except Exception as e:
                messagebox.showerror("오류", f"문서를 여는 중 오류가 발생했습니다:\n{e}")
        else:
            messagebox.showinfo("안내", "현재 HTML 뷰어는 워드(.docx) 파일만 지원합니다.")

    def open_current_document(self):
        if hasattr(self, 'current_filepath') and self.current_filepath:
            try:
                import os
                os.startfile(self.current_filepath)
                messagebox.showinfo("안내", "워드(Word) 프로그램으로 원본 문서를 열었습니다.\n워드에서 직접 표, 서식, 글자 등을 수정한 뒤 [저장] 하시면 됩니다.")
            except Exception as e:
                messagebox.showerror("오류", f"워드를 실행하는 중 오류가 발생했습니다:\n{e}")

    def search_in_viewer(self):
        if not self.html_viewer: return
        query = self.entry_viewer_search.get().strip()
        if not query:
            self.reset_viewer_search()
            return
            
        import re
        safe_query = re.escape(query)
        
        if query != self.last_search_query:
            self.last_search_query = query
            self.current_search_index = 1
        else:
            self.current_search_index += 1
            
        try:
            matches_count = self.html_viewer.find_text(safe_query, select=self.current_search_index, ignore_case=True, highlight_all=True)
            
            if matches_count == 0:
                self.lbl_viewer_search_result.config(text="결과 없음", foreground="red")
            else:
                if self.current_search_index > matches_count:
                    self.current_search_index = 1
                    self.html_viewer.find_text(safe_query, select=self.current_search_index, ignore_case=True, highlight_all=True)
                
                self.lbl_viewer_search_result.config(text=f"{self.current_search_index} / {matches_count} 개", foreground="blue")
        except Exception:
            self.lbl_viewer_search_result.config(text="검색 오류", foreground="red")

    def reset_viewer_search(self):
        self.entry_viewer_search.delete(0, tk.END)
        self.last_search_query = ""
        self.current_search_index = 0
        self.lbl_viewer_search_result.config(text="")
        if self.html_viewer:
            self.html_viewer.find_text("")

    def extract_references(self):
        if not hasattr(self, 'current_doc_text') or not self.current_doc_text:
            messagebox.showwarning("경고", "먼저 문서를 열어주세요.")
            return
            
        import re
        full_text = self.current_doc_text
        
        # 문서 전체를 줄 단위로 스캔하여 규격을 찾습니다. (목차에 의한 잘림 현상 방지)
        lines = full_text.split('\n')
        extracted_count = 0
        found_codes = set()
        
        for i, line in enumerate(lines):
            line = line.strip()
            if not line or len(line) < 4: continue
            
            # 보다 정교한 규격 패턴 매칭 (ASME PAUT 등 복잡한 규격명 완벽 지원 및 부분 단어 오탐지 방지)
            pattern = r'\b(ASME\s+(?:Sec(?:tion|\.)?\s*[IVX]+(?:\s*Art(?:icle|\.)?\s*\d+(?:\s*(?:SE|SA|SB)-[\w\d\-]+)?)?(?:\s*Mandatory\s*Appendix\s*[IVX]+)?|B\s*\d+(?:\.\d+)?)|API\s*(?:Std|Spec)?\s*\d+[A-Z]?|AWS\s*[A-Z]\d+(?:\.\d+)?|ISO\s*\d+(?:-\d+)?|KS\s*[A-Z]\s*\d+(?:-\d+)?|ASTM\s*[A-Z]\s*\d+|EN\s*\d+(?:-\d+)?|ASNT\s*SNT-TC-1A(?:(?:\s*\(\d{4}\s*Ed\.\))?)|KEPIC\s*[A-Z\d]+)'
            match = re.search(pattern, line, re.IGNORECASE)
            
            if match:
                code_name = match.group(1).strip()
                # 불필요한 공백 제거로 정규화 (예: ASME Sec. V -> ASME Sec.V)
                code_name_norm = " ".join(code_name.split())
                
                if code_name_norm.lower() in found_codes:
                    continue
                found_codes.add(code_name_norm.lower())
                
                # 규격 원문 내용 (보통 해당 줄이 제목임)
                details = f"[규격 명칭]\n{line}"
                if i + 1 < len(lines) and lines[i+1].strip() and not re.search(r'^\d+\.', lines[i+1]):
                    details += " " + lines[i+1].strip()
                    
                # 문서 내에서 이 코드가 사용된 부분 검색 (상세내용에 추가)
                usages = []
                # 문서 본문 전체에서 검색 (Reference 목차 제외를 위해 ref_blocks[0] 등 활용 가능하나 단순 정규식 사용)
                for usage_match in re.finditer(r'.{0,40}' + re.escape(code_name) + r'.{0,40}', full_text, re.IGNORECASE):
                    usage_text = usage_match.group(0).replace('\n', ' ').strip()
                    # 너무 중복되지 않도록 필터링
                    if usage_text not in usages and "Reference" not in usage_text:
                        usages.append(usage_text)
                
                if usages:
                    details += "\n\n[문서 내 검색된 사용 예시]\n"
                    # 최대 5개의 사용처만 추가
                    for u in usages[:5]:
                        details += f"- ...{u}...\n"
                        
                # 중복 체크 후 추가
                exists = any(code_name_norm.lower() in d.get("find", "").lower() for d in self.data)
                if not exists:
                    # 규격 접두사를 기반으로 분류(Category) 자동 지정
                    prefix = code_name_norm.split()[0].upper()
                    if prefix in ["ASME", "API", "AWS", "ISO", "KS", "ASTM", "EN", "KEPIC", "ASNT"]:
                        category_name = f"{prefix} 규격"
                    else:
                        category_name = "규격 (자동추출)"
                        
                    self.data.append({
                        "category": category_name,
                        "find": code_name_norm,
                        "replace": "",
                        "details": details
                    })
                    extracted_count += 1
                    
        if extracted_count > 0:
            self.save_data()
            self.refresh_list()
            messagebox.showinfo("추출 완료", f"성공적으로 {extracted_count}개의 새로운 규격을 추출했습니다!\n문서 내 사용 예시도 함께 상세내용에 등록되었습니다.")
            self.notebook.select(self.tab_code) # 결과 확인을 위해 코드 관리 탭으로 자동 이동
        else:
            if found_codes:
                messagebox.showinfo("추출 완료", f"문서에서 {len(found_codes)}개의 규격을 찾았으나, 모두 이미 등록되어 있는 코드입니다.")
            else:
                messagebox.showinfo("추출 실패", "문서에서 규격 코드(ASME, KS, ISO 등)를 찾지 못했습니다.\n본문 텍스트가 추출되지 않았을 수 있습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = CodebookApp(root)
    root.mainloop()
