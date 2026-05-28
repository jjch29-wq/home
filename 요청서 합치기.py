import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import threading
import re

# ============================================================
# Excel Smart Merger v2.7 - FINAL STABLE VERSION
# ============================================================
print("====================================================")
print("  LOADING EXCEL SMART MERGER v2.7 (FINAL FIX)      ")
print("====================================================")

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Smart Merger - 지능형 요청서 합치기 v2.7")
        self.root.geometry("850x850")
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

        title_label = tk.Label(main_frame, text="🧠 지능형 엑셀 병합 도구 v2.7", 
                               font=("Malgun Gothic", 18, "bold"), fg="#0be881", bg="#2c3e50")
        title_label.pack(pady=(0, 20))

        # 폴더 선택
        folder_frame = tk.Frame(main_frame, bg="#34495e", padx=10, pady=10)
        folder_frame.pack(fill=tk.X, pady=5)
        self.folder_path_var = tk.StringVar(value="폴더를 선택해주세요...")
        tk.Label(folder_frame, textvariable=self.folder_path_var, fg="#bdc3c7", bg="#34495e", 
                 font=("Malgun Gothic", 10), anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(folder_frame, text="📄 파일 선택", command=self.select_files, bg="#9b59b6", fg="white", 
                  font=("Malgun Gothic", 10, "bold"), relief=tk.FLAT, padx=15).pack(side=tk.RIGHT, padx=(5, 0))
        tk.Button(folder_frame, text="📁 폴더 선택", command=self.select_folder, bg="#3498db", fg="white", 
                  font=("Malgun Gothic", 10, "bold"), relief=tk.FLAT, padx=15).pack(side=tk.RIGHT)

        # 추출 설정
        smart_frame = tk.LabelFrame(main_frame, text=" 추출 설정 (Keywords) ", 
                                    font=("Malgun Gothic", 10, "bold"), fg="#3498db", bg="#2c3e50", padx=10, pady=10)
        smart_frame.pack(fill=tk.X, pady=15)
        
        tk.Label(smart_frame, text="추출할 키워드 (쉼표로 구분):", font=("Malgun Gothic", 9), fg="#ecf0f1", bg="#2c3e50").pack(anchor="w")
        
        self.keyword_var = tk.StringVar(value="No, Joint, Dwg, Size, THK, Result, Date, Report No, Identification No")
        keyword_entry = tk.Entry(smart_frame, textvariable=self.keyword_var, font=("Malgun Gothic", 10), bg="#ecf0f1")
        keyword_entry.pack(fill=tk.X, pady=5)
        
        self.only_totals_var = tk.BooleanVar(value=False)
        tk.Checkbutton(smart_frame, text="합계(소계/총합계)만 추출", variable=self.only_totals_var,
                       font=("Malgun Gothic", 9), fg="#ecf0f1", bg="#2c3e50", selectcolor="#2c3e50",
                       activebackground="#2c3e50", activeforeground="#ecf0f1").pack(anchor="w", pady=(2, 5))
        
        tk.Label(smart_frame, text="💡 v2.7: 중복 헤더 방지 및 전역 정보(Report No) 추출이 최적화되었습니다.", 
                 font=("Malgun Gothic", 8), fg="#95a5a6", bg="#2c3e50").pack(anchor="w")

        # 실시간 로그창
        log_label = tk.Label(main_frame, text="작업 진행 로그:", font=("Malgun Gothic", 10), fg="#ecf0f1", bg="#2c3e50")
        log_label.pack(anchor="w")
        self.log_text = tk.Text(main_frame, height=22, bg="#1e272e", fg="#0be881", font=("Consolas", 9), padx=10, pady=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.status_var = tk.StringVar(value="대기 중...")
        tk.Label(main_frame, textvariable=self.status_var, font=("Malgun Gothic", 10, "bold"), fg="#e67e22", bg="#2c3e50").pack(pady=5)

        self.btn_merge = tk.Button(main_frame, text="🚀 지능형 병합 시작 (v2.7)", command=self.start_merge_thread,
                                   bg="#2ecc71", fg="white", font=("Malgun Gothic", 12, "bold"), relief=tk.FLAT, pady=10)
        self.btn_merge.pack(fill=tk.X, pady=10)
        self.btn_merge["state"] = tk.DISABLED

    def add_log(self, msg):
        timestamp = datetime.now().strftime("[%H:%M:%S] ")
        self.log_text.insert(tk.END, timestamp + msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_folder = folder
            self.folder_path_var.set(folder)
            self.scan_files()

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="병합할 엑셀 파일들을 선택하세요",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
        )
        if files:
            # 선택된 파일들의 폴더를 저장 위치로 사용
            self.selected_folder = os.path.dirname(files[0])
            self.folder_path_var.set(f"개별 파일 선택됨 ({len(files)}개)")
            
            # 파일명만 추출하여 리스트에 저장 (기존 병합 로직 호환성 유지)
            self.excel_files = [os.path.basename(f) for f in files if "Smart_Merged" not in f and not os.path.basename(f).startswith('~$')]
            
            if self.excel_files:
                self.btn_merge["state"] = tk.NORMAL
                self.status_var.set(f"총 {len(self.excel_files)}개 파일 준비 완료")
                self.add_log(f"개별 파일 선택 완료: {len(self.excel_files)}개")
            else:
                self.add_log("⚠️ 유효한 엑셀 파일이 없습니다.")

    def scan_files(self):
        # 작업 결과물(Smart_Merged)은 다시 병합하지 않도록 제외
        self.excel_files = [f for f in os.listdir(self.selected_folder) 
                            if (f.endswith('.xlsx') or f.endswith('.xlsm')) and not f.startswith('~$') and "Smart_Merged" not in f]
        if self.excel_files:
            self.btn_merge["state"] = tk.NORMAL
            self.status_var.set(f"총 {len(self.excel_files)}개 파일 준비 완료")
            self.add_log(f"파일 검색 완료: {len(self.excel_files)}개")
        else:
            self.add_log("⚠️ 폴더에 엑셀 파일이 없습니다.")

    def start_merge_thread(self):
        self.btn_merge["state"] = tk.DISABLED
        self.log_text.delete("1.0", tk.END)
        threading.Thread(target=self.merge_logic, daemon=True).start()

    def normalize(self, text):
        if pd.isna(text): return ""
        t = str(text).lower()
        t = re.sub(r'[^a-z0-9가-힣]', '', t)
        return t.strip()

    def merge_logic(self):
        try:
            all_data = []
            raw_kws = [k.strip() for k in self.keyword_var.get().split(',') if k.strip()]
            norm_keywords = [self.normalize(k) for k in raw_kws]
            
            if not norm_keywords:
                messagebox.showwarning("알림", "검색할 키워드를 입력해주세요.")
                return

            self.add_log("🚀 v2.8 지능형 병합 엔진 기동 중...")

            # 1. 유의어(Synonyms) 사전 정의
            SYNONYMS = {
                "no": ["no", "no.", "seq", "순번", "연번", "일련번호", "번호", "num", "idx", "호"],
                "dwg": ["dwg", "dwgno", "dwg.no", "도면", "도면번호", "도면명", "drawing", "drawingno", "drawing.no", "iso"],
                "joint": ["joint", "jointno", "joint.no", "조인트", "jnt", "jntno", "point", "포인트"],
                "size": ["size", "규격", "구경", "사이즈", "dia", "nps", "pipe size", "파이프 규격"],
                "thk": ["thk", "thickness", "두께", "t", "thk.", "thick"],
                "result": ["result", "결과", "판정", "결과판정", "decision", "판정결과"],
                "date": ["date", "날짜", "검사일", "검사일자", "일자", "dateofexam", "요청일", "요청일자"],
                "reportno": ["reportno", "report.no", "report_no", "성적서번호", "성적서", "보고서번호", "보고서"],
                "identificationno": ["identificationno", "idno", "id_no", "관리번호", "식별번호", "id"]
            }
            
            # 정규화된 유의어 맵 생성
            NORM_SYNONYMS = {k: [self.normalize(s) for s in v] for k, v in SYNONYMS.items()}
            
            # 제목 줄 찾기 점수 계산용 핵심 키워드 리스트 (성적서번호 등 전역 메타 정보 제외하여 갑지 오진 차단)
            header_kws = ['no', 'joint', 'dwg', 'size', 'thk', 'result', 'date']

            for file in self.excel_files:
                self.add_log(f"--- 분석: {file} ---")
                file_path = os.path.join(self.selected_folder, file)
                
                try:
                    xls = pd.ExcelFile(file_path)
                    
                    for sheet_name in xls.sheet_names:
                        raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                        
                        # 1. 전역 메타데이터 (Report No) 추출
                        meta_info = {}
                        for r_idx, row in raw_df.iterrows():
                            if r_idx > 50: break
                            row_vals = row.values
                            for c_idx, val in enumerate(row_vals):
                                if pd.isna(val): continue
                                s_val = str(val)
                                s_upper = s_val.upper()
                                if ("REPORT" in s_upper and "NO" in s_upper) or ("성적서" in s_val and "번호" in s_val):
                                    extracted_val = ""
                                    if ":" in s_val:
                                        extracted_val = s_val.split(":", 1)[1].strip()
                                    else:
                                        tmp = re.sub(r'(?i)(report\s*no\.?|성적서\s*번호)', '', s_val).strip()
                                        if tmp: extracted_val = tmp
                                    
                                    if not extracted_val or extracted_val.upper() == "NAN":
                                        for offset in range(1, 4):
                                            if c_idx + offset < len(row_vals):
                                                v = str(row_vals[c_idx + offset]).strip()
                                                if v and v.upper() != "NAN" and len(v) < 30:
                                                    extracted_val = v
                                                    break
                                    
                                    if extracted_val and extracted_val.upper() != "NAN":
                                        clean_val = re.split(r'(?i)\n|/|\s{2,}|date|일자|page|페이지|\(|\[|<', extracted_val)[0].strip()
                                        clean_val = re.split(r'(?i)\s+[가-힣]{2,}|\s+(?:rev|sheet|insp|note|remark)', clean_val)[0].strip()
                                        if ":" in clean_val:
                                            tmp = clean_val.split(":")[0].strip()
                                            clean_val = re.sub(r'\s+[A-Za-z가-힣]+$', '', tmp).strip()
                                        clean_val = re.split(r'(?i)\s+(?:IP|LF|CR|PO|UC|BT|NF|INCOMPLETE|LACK|DEFECT|결함)\b', clean_val)[0].strip()
                                        clean_val = re.sub(r'^[:\-]+|[:\-]+$', '', clean_val).strip()
                                        
                                        if clean_val:
                                            if re.search(r'\d', clean_val) or clean_val.upper() in ['N/A', '-', 'TBD', 'NA']:
                                                meta_info["Report No"] = clean_val
                                                self.add_log(f"   📌 메타데이터 정밀 추출: Report No -> {clean_val}")
                                                break
                                            else:
                                                extracted_val = ""
                            if "Report No" in meta_info: break

                        if "Report No" not in meta_info:
                            base_name = os.path.splitext(file)[0]
                            meta_info["Report No"] = base_name
                            self.add_log(f"   📌 메타데이터 (파일명 기준): Report No -> {base_name}")

                        # 2. 제목 줄 찾기 (유의어 기반 고도화 점수제)
                        best_row = 0
                        max_score = 0
                        for idx, row in raw_df.iterrows():
                            if idx > 60: break
                            row_content = "".join([str(v) for v in row.values if pd.notna(v)])
                            norm_content = self.normalize(row_content)
                            
                            score = 0
                            for kw in header_kws:
                                norm_kw = self.normalize(kw)
                                if norm_kw in norm_content or any(syn in norm_content for syn in NORM_SYNONYMS.get(norm_kw, [])):
                                    score += 1
                                    
                            if score > max_score:
                                max_score = score
                                best_row = idx
                        
                        # 유효한 테이블 감지 조건 (점수 3점 이상)
                        if max_score >= 3:
                            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=best_row)
                            
                            # Multi-row Header 지원
                            if len(df) > 0:
                                row2 = df.iloc[0]
                                row2_content = "".join([str(v) for v in row2.values if pd.notna(v)])
                                if sum(1 for kw in norm_keywords if kw in self.normalize(row2_content)) >= 1:
                                    new_cols = []
                                    for c1, c2 in zip(df.columns, row2):
                                        c1_s = str(c1) if not str(c1).startswith('Unnamed') else ""
                                        c2_s = str(c2) if not str(c2).startswith('nan') else ""
                                        combined = (c1_s + " " + c2_s).strip()
                                        new_cols.append(combined if combined else "Col")
                                    df.columns = new_cols
                                    df = df.iloc[1:]
                            
                            # 컬럼 정리 및 중복 방어
                            df.columns = [str(c).strip() for c in df.columns]
                            df = df.loc[:, ~df.columns.duplicated()]
                            
                            # 최종 컬럼 유의어 기반 정밀 매칭 (1키워드 당 1컬럼 매칭)
                            final_cols = []
                            norm_col_map = {col: self.normalize(col) for col in df.columns}
                            keyword_to_col = {}
                            
                            for raw_kw in raw_kws:
                                kw = self.normalize(raw_kw)
                                # 1. 정확 매칭
                                match = next((orig for orig, norm in norm_col_map.items() if norm == kw and orig not in final_cols), None)
                                if match:
                                    final_cols.append(match)
                                    keyword_to_col[kw] = match
                                    continue
                                
                                # 2. 유의어 매칭
                                syns = NORM_SYNONYMS.get(kw, [])
                                match = next((orig for orig, norm in norm_col_map.items() if norm in syns and orig not in final_cols), None)
                                if match:
                                    final_cols.append(match)
                                    keyword_to_col[kw] = match
                                    continue
                                    
                                # 3. 유사(부분) 매칭 (너무 짧은 'no' 키워드 등이 엉뚱하게 매칭되는 것 방어)
                                if len(kw) >= 3:
                                    match = next((orig for orig, norm in norm_col_map.items() if (kw in norm or any(syn in norm for syn in syns)) and orig not in final_cols), None)
                                    if match:
                                        final_cols.append(match)
                                        keyword_to_col[kw] = match
                                        continue

                            # 핵심 Joint 컬럼이 없으면 본문 데이터가 없는 갑지/요약 시트로 판정하여 스킵
                            if "joint" not in keyword_to_col:
                                self.add_log(f"   ⚠️ 시트 스킵: '{sheet_name}' (Joint 컬럼 유실)")
                                continue
                                
                            # 매칭 컬럼 필터링 및 표준 키워드로 컬럼명 변경 (병합 시 완벽 정렬 보장!)
                            df = df[final_cols]
                            rename_map = {keyword_to_col[kw]: raw_kw for raw_kw in raw_kws if (kw := self.normalize(raw_kw)) in keyword_to_col}
                            df.rename(columns=rename_map, inplace=True)
                            
                            # [NEW] 세로 병합 셀(Merged Cells) 복원을 위한 Forward Fill 적용 (표준화된 컬럼명 기준)
                            if "Dwg" in df.columns:
                                df["Dwg"] = df["Dwg"].ffill()
                            if "Joint" in df.columns:
                                df["Joint"] = df["Joint"].ffill()
                                
                            # [NEW] 데이터 행 필터링 및 서명/꼬리표 일괄 청소 (표준화된 컬럼명 기준)
                            filter_col = "No" if "No" in df.columns else ("Joint" if "Joint" in df.columns else None)
                            
                            if filter_col:
                                df = df.dropna(subset=[filter_col])
                                df = df[df[filter_col].astype(str).str.strip() != ""]
                                
                                if filter_col == "No":
                                    # 순번 필터: 숫자를 하나 이상 포함하되 한글/영문 꼬리표는 배제
                                    df = df[df["No"].astype(str).str.contains(r'\d', regex=True, na=False)]
                                    df = df[~df["No"].astype(str).str.contains(r'[a-zA-Z가-힣]', regex=True, na=False)]
                                else:
                                    # 조인트 필터: 하단 서명란이나 빈 줄 꼬리표 제거
                                    exclude_terms = r'(?i)page|페이지|total|합계|sub-total|소계|grand|seoul|inspection|testing|report|project|client|customer|date'
                                    df = df[~df[filter_col].astype(str).str.contains(exclude_terms, regex=True, na=False)]
                                
                                self.add_log(f"   🎯 테이블 감지됨: '{sheet_name}' 시트, {best_row+1}행 -> {len(df)}행 추출 완료")
                            
                            # 메타데이터 주입
                            for m_key, m_val in meta_info.items():
                                df[m_key] = m_val
                                
                            all_data.append(df)
                            
                except Exception as fe:
                    self.add_log(f"   ❌ 시트 분석 중 오류 발생: {str(fe)}")

            if not all_data:
                self.add_log("❌ 병합할 데이터가 없습니다.")
                return

            self.add_log("📊 전체 데이터 병합 및 단일 매칭 필터링 중...")
            combined_df = pd.concat(all_data, ignore_index=True, sort=False)
            combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
            
            # 최종 중복 데이터 제거
            combined_df.drop_duplicates(inplace=True)
            
            # [NEW] No. of Film 수량 원본 파일별 소계 및 전체 총합산 기능
            film_col = next((c for c in combined_df.columns if "film" in self.normalize(c) and "size" not in self.normalize(c) and "loc" not in self.normalize(c)), None)
            joint_col = next((c for c in combined_df.columns if self.normalize(c) == "joint"), None)
            
            if len(combined_df) > 0:
                no_col_name = next((c for c in combined_df.columns if self.normalize(c) == "no"), None)
                
                report_col = next((c for c in combined_df.columns if self.normalize(c) == "reportno" or any(syn in self.normalize(c) for syn in NORM_SYNONYMS.get("reportno", []))), None)
                label_col = no_col_name if no_col_name else next((c for c in combined_df.columns if c != joint_col and c != film_col and c != report_col), combined_df.columns[0])
                
                if film_col:
                    combined_df['_temp_numeric_film'] = pd.to_numeric(combined_df[film_col].astype(str).str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
                
                new_dfs = []
                grand_total = 0
                grand_joint_total = 0
                
                if report_col:
                    for rep_no, group in combined_df.groupby(report_col, sort=False):
                        new_dfs.append(group)
                        
                        sub_total = group['_temp_numeric_film'].sum() if film_col else 0
                        grand_total += sub_total
                        
                        sub_joint = 0
                        if joint_col:
                            valid_joints = group[joint_col].dropna().astype(str).str.strip()
                            valid_joints = valid_joints[valid_joints != ""]
                            sub_joint = valid_joints.nunique()
                        grand_joint_total += sub_joint
                        
                        sub_row = {col: "" for col in combined_df.columns}
                        if label_col == report_col:
                            sub_row[report_col] = f"{rep_no} (소계)"
                        else:
                            sub_row[label_col] = "Sub-Total"
                            sub_row[report_col] = rep_no
                        if film_col and film_col != label_col:
                            sub_row[film_col] = int(sub_total) if sub_total % 1 == 0 else sub_total
                        if joint_col and joint_col != label_col:
                            sub_row[joint_col] = int(sub_joint)
                        new_dfs.append(pd.DataFrame([sub_row]))
                    
                    combined_df = pd.concat(new_dfs, ignore_index=True)
                else:
                    grand_total = combined_df['_temp_numeric_film'].sum() if film_col else 0
                    if joint_col:
                        valid_joints = combined_df[joint_col].dropna().astype(str).str.strip()
                        valid_joints = valid_joints[valid_joints != ""]
                        grand_joint_total = valid_joints.nunique()
                    else:
                        grand_joint_total = 0

                total_row = {col: "" for col in combined_df.columns}
                if label_col == report_col:
                    total_row[report_col] = "총합계 (Grand Total)"
                else:
                    total_row[label_col] = "Grand Total"
                if film_col and film_col != label_col:
                    total_row[film_col] = int(grand_total) if grand_total % 1 == 0 else grand_total
                if joint_col and joint_col != label_col:
                    total_row[joint_col] = int(grand_joint_total)
                combined_df = pd.concat([combined_df, pd.DataFrame([total_row])], ignore_index=True)
                
                log_msg = "   ➕ 소계 및 총합계 추가 완료"
                if film_col or joint_col:
                    log_msg += " (수량 합산 포함)"
                self.add_log(log_msg)
                
                if film_col and '_temp_numeric_film' in combined_df.columns:
                    combined_df.drop(columns=['_temp_numeric_film'], inplace=True)
                
                if self.only_totals_var.get():
                    mask = combined_df.astype(str).apply(lambda row: row.str.contains("Sub-Total|Grand Total|소계|총합계", case=False).any(), axis=1)
                    combined_df = combined_df[mask]

            out_name = f"Final_Smart_Merged_v2.8_{datetime.now().strftime('%H%M%S')}.xlsx"
            out_path = os.path.join(self.selected_folder, out_name)
            combined_df.to_excel(out_path, index=False)
            
            self.add_log(f"✨ 완료: {out_name}")
            self.status_var.set(f"✅ 저장 완료: {out_name}")
            messagebox.showinfo("성공", f"지능형 병합이 성공적으로 완료되었습니다!\n파일명: {out_name}")
            
        except Exception as e:
            self.add_log(f"❌ 치명적 오류: {str(e)}")
            messagebox.showerror("오류", f"프로세스 오류: {e}")
        finally:
            self.btn_merge["state"] = tk.NORMAL

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
