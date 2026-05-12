import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import threading

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Smart Merger - 지능형 요청서 합치기")
        self.root.geometry("700x650")
        self.root.configure(bg="#2c3e50")
        
        self.selected_folder = ""
        self.excel_files = []
        
        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Main.TFrame", background="#2c3e50")
        
        main_frame = ttk.Frame(self.root, style="Main.TFrame", padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = tk.Label(main_frame, text="🧠 지능형 엑셀 병합 도구", 
                               font=("Malgun Gothic", 18, "bold"), fg="#ecf0f1", bg="#2c3e50")
        title_label.pack(pady=(0, 20))

        # 폴더 선택
        folder_frame = tk.Frame(main_frame, bg="#34495e", padx=10, pady=10)
        folder_frame.pack(fill=tk.X, pady=5)
        self.folder_path_var = tk.StringVar(value="폴더를 선택해주세요...")
        tk.Label(folder_frame, textvariable=self.folder_path_var, fg="#bdc3c7", bg="#34495e", 
                 font=("Malgun Gothic", 10), anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(folder_frame, text="📁 폴더 선택", command=self.select_folder, bg="#3498db", fg="white", 
                  font=("Malgun Gothic", 10, "bold"), relief=tk.FLAT, padx=15).pack(side=tk.RIGHT)

        # [NEW] 지능형 헤더 검색 설정
        smart_frame = tk.LabelFrame(main_frame, text=" 지능형 제목 줄 찾기 설정 ", 
                                    font=("Malgun Gothic", 10, "bold"), fg="#3498db", bg="#2c3e50", padx=10, pady=10)
        smart_frame.pack(fill=tk.X, pady=15)
        
        tk.Label(smart_frame, text="찾을 키워드 (쉼표로 구분):", font=("Malgun Gothic", 9), fg="#ecf0f1", bg="#2c3e50").pack(anchor="w")
        
        self.keyword_var = tk.StringVar(value="번호, 품명, 항목, 날짜, 일자")
        keyword_entry = tk.Entry(smart_frame, textvariable=self.keyword_var, font=("Malgun Gothic", 10), bg="#ecf0f1")
        keyword_entry.pack(fill=tk.X, pady=5)
        
        tk.Label(smart_frame, text="💡 입력한 단어가 포함된 줄을 자동으로 찾아 제목으로 사용합니다.", 
                 font=("Malgun Gothic", 8), fg="#95a5a6", bg="#2c3e50").pack(anchor="w")

        # 파일 리스트
        list_label = tk.Label(main_frame, text="발견된 파일 목록:", font=("Malgun Gothic", 10), fg="#ecf0f1", bg="#2c3e50")
        list_label.pack(anchor="w")
        self.file_listbox = tk.Listbox(main_frame, bg="#ecf0f1", font=("Consolas", 10))
        self.file_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

        self.status_var = tk.StringVar(value="대기 중...")
        tk.Label(main_frame, textvariable=self.status_var, font=("Malgun Gothic", 9), fg="#e67e22", bg="#2c3e50").pack(pady=5)

        self.btn_merge = tk.Button(main_frame, text="🚀 지능형 병합 시작", command=self.start_merge_thread,
                                   bg="#2ecc71", fg="white", font=("Malgun Gothic", 12, "bold"), relief=tk.FLAT, pady=10)
        self.btn_merge.pack(fill=tk.X, pady=10)
        self.btn_merge["state"] = tk.DISABLED

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_folder = folder
            self.folder_path_var.set(folder)
            self.scan_files()

    def scan_files(self):
        self.file_listbox.delete(0, tk.END)
        self.excel_files = [f for f in os.listdir(self.selected_folder) if f.endswith('.xlsx') and not f.startswith('~$')]
        for f in self.excel_files: self.file_listbox.insert(tk.END, f"  📄 {f}")
        if self.excel_files:
            self.btn_merge["state"] = tk.NORMAL
            self.status_var.set(f"총 {len(self.excel_files)}개 파일 준비 완료")

    def start_merge_thread(self):
        self.btn_merge["state"] = tk.DISABLED
        threading.Thread(target=self.merge_logic, daemon=True).start()

    def merge_logic(self):
        try:
            all_data = []
            keywords = [k.strip() for k in self.keyword_var.get().split(',') if k.strip()]
            
            for file in self.excel_files:
                self.status_var.set(f"검색 중: {file}")
                file_path = os.path.join(self.selected_folder, file)
                
                # 1. 헤더 없이 일단 읽기
                raw_df = pd.read_excel(file_path, header=None)
                
                header_row_index = 0
                found = False
                
                # 2. 행을 돌며 키워드가 포함된 행 찾기 (상위 20행까지만 검색)
                for idx, row in raw_df.iterrows():
                    if idx > 20: break 
                    # 모든 셀 값을 문자열로 변환하여 합치기 (오류 방지)
                    row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
                    if any(kw in row_str for kw in keywords):
                        header_row_index = idx
                        found = True
                        break
                
                # 3. 찾은 행을 헤더로 다시 읽기
                if found:
                    df = pd.read_excel(file_path, skiprows=header_row_index)
                else:
                    df = pd.read_excel(file_path) # 못 찾으면 기본값(0행)
                
                # 컬럼명 정리 (공백 제거 등)
                df.columns = [str(c).strip() for c in df.columns]
                all_data.append(df)

            # 병합
            combined_df = pd.concat(all_data, ignore_index=True)
            
            # [NEW] 키워드에 해당하는 열만 추출
            filtered_cols = []
            for kw in keywords:
                # 해당 키워드가 포함된 열 이름을 모두 찾음
                match_cols = [col for col in combined_df.columns if kw in col]
                filtered_cols.extend(match_cols)
            
            # 중복된 컬럼명 제거 (순서 유지)
            seen = set()
            filtered_cols = [x for x in filtered_cols if not (x in seen or seen.add(x))]

            if filtered_cols:
                combined_df = combined_df[filtered_cols]
                self.status_var.set(f"✅ 키워드 필터링 완료 ({len(filtered_cols)}개 항목)")
            else:
                self.status_var.set("⚠️ 키워드와 일치하는 열을 찾지 못해 전체를 표시합니다.")

            # 중복 행 제거
            combined_df.drop_duplicates(inplace=True)

            
            # 저장
            out_name = f"지능형_병합_{datetime.now().strftime('%H%M%S')}.xlsx"
            combined_df.to_excel(os.path.join(self.selected_folder, out_name), index=False)
            
            self.status_var.set(f"✅ 완료: {out_name}")
            messagebox.showinfo("성공", f"키워드를 기준으로 제목을 찾아 병합했습니다.\n파일명: {out_name}")
            
        except Exception as e:
            messagebox.showerror("오류", str(e))
        finally:
            self.btn_merge["state"] = tk.NORMAL

if __name__ == "__main__":
    root = tk.Tk(); app = ExcelMergerApp(root); root.mainloop()
