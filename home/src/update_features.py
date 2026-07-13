import re
import os

with open('c:/Users/jjch2/Desktop/PMI/home/src/코드절차서관리.py', encoding='utf-8') as f:
    code = f.read()

# 1. Update load_document to use _load_document_by_path and add btn_apply_db
load_doc_code = '''
    def load_document(self):
        if self.html_viewer is None or mammoth is None:
            messagebox.showerror("오류", "문서를 렌더링하기 위한 필수 라이브러리가 없습니다.")
            return
            
        filepath = filedialog.askopenfilename(
            title="절차서 문서 열기 (Word)",
            filetypes=[("워드 파일", "*.docx"), ("모든 파일", "*.*")]
        )
        if filepath:
            self._load_document_by_path(filepath)

    def _load_document_by_path(self, filepath):
        if filepath.lower().endswith('.docx'):
            try:
                with open(filepath, "rb") as docx_file:
                    result = mammoth.convert_to_html(docx_file)
                    html = result.value
                    
                    docx_file.seek(0)
                    text_result = mammoth.extract_raw_text(docx_file)
                    self.current_doc_text = text_result.value
                    
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
                
                self.current_filepath = filepath
                self.btn_edit_doc.config(state="normal")
                self.btn_apply_db.config(state="normal")
            except Exception as e:
                messagebox.showerror("오류", f"문서를 여는 중 오류가 발생했습니다:\\n{e}")
        else:
            messagebox.showinfo("안내", "현재 HTML 뷰어는 워드(.docx) 파일만 지원합니다.")

    def apply_db_to_current(self):
        if not hasattr(self, 'current_filepath') or not self.current_filepath:
            messagebox.showwarning("경고", "먼저 문서를 열어주세요.")
            return
            
        rules = [(d["find"], d["replace"]) for d in self.data if d.get("find") and d.get("replace")]
        if not rules:
            messagebox.showwarning("경고", "코드 DB에 바꿀 내용(Replace)이 설정된 규격이 하나도 없습니다.")
            return
            
        if not messagebox.askyesno("일괄 적용 확인", f"현재 열려있는 문서에 코드 DB의 변환 규칙 {len(rules)}개를 모두 적용하시겠습니까?\\n(원본 파일은 '_수정본' 이라는 이름으로 같은 폴더에 안전하게 저장됩니다.)"):
            return
            
        try:
            dir_name = os.path.dirname(self.current_filepath)
            base_name = os.path.basename(self.current_filepath)
            name, ext = os.path.splitext(base_name)
            output_filepath = os.path.join(dir_name, f"{name}_수정본{ext}")
            
            self.process_docx(self.current_filepath, output_filepath, rules)
            
            self._load_document_by_path(output_filepath)
            messagebox.showinfo("적용 완료", f"현재 문서에 {len(rules)}개의 변환 규칙을 적용하고 뷰어를 새로고침했습니다!\\n\\n저장 경로: {output_filepath}")
        except Exception as e:
            messagebox.showerror("오류", f"문서 일괄 변환 중 오류가 발생했습니다:\\n{e}")
'''

code = re.sub(r'    def load_document\(self\):.*?else:\n\s+messagebox.showinfo\("안내", "현재 HTML 뷰어는 워드\(\.docx\) 파일만 지원합니다\."\)', load_doc_code.strip(), code, flags=re.DOTALL)

# Add btn_apply_db to create_viewer_widgets
btn_apply_code = '''
        self.btn_edit_doc = ttk.Button(ctrl_frame, text="📝 원본 워드로 열어서 직접 수정하기", command=self.open_current_document, state="disabled")
        self.btn_edit_doc.pack(side="left", padx=10)
        
        self.btn_apply_db = ttk.Button(ctrl_frame, text="⚡ 현재 문서에 코드 DB 일괄 적용", command=self.apply_db_to_current, state="disabled")
        self.btn_apply_db.pack(side="left", padx=10)
'''
code = code.replace('''        self.btn_edit_doc = ttk.Button(ctrl_frame, text="📝 원본 워드로 열어서 직접 수정하기", command=self.open_current_document, state="disabled")
        self.btn_edit_doc.pack(side="left", padx=10)''', btn_apply_code.strip())

# Add load_from_code_db to create_batch_widgets
preset_btns = '''
        ttk.Button(preset_frame, text="💾 현재 목록 저장", command=self.save_preset).pack(side="right", padx=2)
        ttk.Button(preset_frame, text="📂 목록 불러오기", command=self.load_preset).pack(side="right", padx=2)
        ttk.Button(preset_frame, text="📚 코드 DB에서 최신 규격 끌어오기", command=self.load_from_code_db).pack(side="right", padx=10)
'''
code = code.replace('''        ttk.Button(preset_frame, text="💾 현재 목록 저장", command=self.save_preset).pack(side="right", padx=2)
        ttk.Button(preset_frame, text="📂 목록 불러오기", command=self.load_preset).pack(side="right", padx=2)''', preset_btns.strip())

# Add load_from_code_db method
load_method = '''
    def load_from_code_db(self):
        rules = [(d["find"], d["replace"]) for d in self.data if d.get("find") and d.get("replace")]
        if not rules:
            messagebox.showwarning("경고", "코드 관리 DB에 '바꿀 내용'이 설정된 규격이 없습니다.")
            return
            
        if messagebox.askyesno("불러오기 확인", f"코드 DB에 저장된 {len(rules)}개의 변환 규칙을 목록에 가져오시겠습니까?\\n(기존 목록은 초기화됩니다.)"):
            for item in self.batch_tree.get_children():
                self.batch_tree.delete(item)
            for f_text, r_text in rules:
                self.batch_tree.insert("", "end", values=(f_text, r_text))
            messagebox.showinfo("불러오기 완료", "성공적으로 코드 DB에서 규칙을 불러왔습니다!")

    def load_preset(self):
'''
code = code.replace('    def load_preset(self):', load_method.strip())

with open('c:/Users/jjch2/Desktop/PMI/home/src/코드절차서관리.py', 'w', encoding='utf-8') as f:
    f.write(code)
print('Done')
