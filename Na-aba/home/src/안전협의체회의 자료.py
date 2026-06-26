import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import sys

# 상위 폴더(부모 디렉토리)에 있는 모듈을 불러오기 위해 경로 추가
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from create_excel_form import generate_excel
from fill_hwp_form import generate_hwp

# 설정 파일 경로
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "form_config.json")

# 기본 데이터 템플릿
DEFAULT_DATA = {
    "템플릿_경로": r"C:\Users\-\OneDrive\바탕 화면\1. 안전보건협의체 수급업체 회의자료(서식).hwp",
    "제출년월": "2026년 7월",
    "실적년월": "2026년 6월",
    "계획년월": "2026년 7월",
    "계약명": "가산~가평 천연가스 공급시설 비파괴검사 용역",
    "계약기간": "2024.01.01 ~ 2026.12.31",
    "계약상대자(업체명)": "서울검사(주)",
    "현장대리인": "곽재운",
    "작업의 시작시간": "08:00",
    "작업 또는 작업장 간의 연락방법": "무전기 및 휴대전화",
    "재해발생 위험시의 대피방법": "작업 중지 후 비상연락망 전파 및 집결지 대피",
    "사업자와 수급인 또는 수급인 상호간의 연락방법": "현장소장 및 안전관리자 핫라인 구축",
    "작업공정의 조정 및 협의 요청사항": "타 공정 간섭구간 작업 전 사전 협의 요망",
    "주요 활동실적": "주배관 RT 검사 500매 완료",
    "주요 활동계획": "주배관 및 관리소 RT/PT 검사 예정",
    "최초위험성평가_실시여부": "",
    "최초위험성평가_작성날짜": "",
    "정기위험성평가_실시여부": "",
    "정기위험성평가_작성날짜": "",
    "수시위험성평가_실시여부": "O",
    "수시위험성평가_작성날짜": "2026.06.29",
    "관리주관부서": "안전환경부",
    "장소": "가산~가평 주배관 공사 현장",
    "위험성평가_중점관리항목": "가스 누출 방지 및 용접부 비파괴검사 철저",
    "위험성평가_중점관리항목": "가스 누출 방지 및 용접부 비파괴검사 철저",
    "조치사진_경로": "",
    "개선후사진_경로": "",
    "감소대책_위험성제거": False,
    "감소대책_공학적": False,
    "감소대책_관리적": False,
    "감소대책_개인보호구": False,
    "개선조치_이행사항": "",
    "아차사고_사고명": "",
    "아차사고_발생일시": "",
    "아차사고_장소": "",
    "아차사고_보고자": "",
    "아차사고_소속": "",
    "아차사고_사고내용": "",
    "아차사고_원인분석": "",
    "아차사고_조치전사진": "",
    "아차사고_조치후사진": "",
    "건의_개진사항": "",
    "건의_제안사유": "",
    "교육관리_수급업체명": "서울검사(주)",
    "교육관리_계약명": "가산~가평 천연가스 공급시설 비파괴검사 용역",
    "교육관리_현장대리인": "곽재운",
    "교육관리_년월": "2026년 6월",
    "교육관리_직원목록": []
}

class FormGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("안전보건협의체 회의자료 자동 완성기")
        self.root.geometry("650x750")
        
        self.data = self.load_config()
        self.entries = {}
        self.edu_records = list(self.data.get("교육관리_직원목록", []))
        
        self.create_widgets()
        
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 병합하여 새로운 키가 있으면 기본값 사용
                    merged = DEFAULT_DATA.copy()
                    merged.update(data)
                    return merged
            except Exception as e:
                print(f"설정 파일 로드 실패: {e}")
        return DEFAULT_DATA.copy()
        
    def save_config(self):
        current_data = self.get_current_data()
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(current_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"설정 파일 저장 실패: {e}")

    def get_current_data(self):
        current_data = {}
        for key, widget in self.entries.items():
            if isinstance(widget, tk.Text):
                current_data[key] = widget.get("1.0", "end-1c")
            elif isinstance(widget, tk.BooleanVar):
                current_data[key] = widget.get()
            else:
                current_data[key] = widget.get()
        current_data["교육관리_직원목록"] = self.edu_records
        return current_data

    def create_widgets(self):
        # 타이틀 (최상단 고정)
        title_lbl = ttk.Label(self.root, text="회의자료 자동 완성기", font=("맑은 고딕", 16, "bold"))
        title_lbl.pack(pady=(15, 5), padx=20, anchor="w")

        # 0. 템플릿 경로 설정 (최상단 고정)
        lbl_frame0 = ttk.LabelFrame(self.root, text="템플릿 파일 설정")
        lbl_frame0.pack(fill="x", padx=20, pady=5)
        
        frame0 = ttk.Frame(lbl_frame0)
        frame0.pack(fill="x", padx=20, pady=5)
        
        lbl0 = ttk.Label(frame0, text="원본 한글 양식(.hwp):", width=20, anchor="e")
        lbl0.pack(side="left", padx=(0, 10))
        
        self.template_entry = ttk.Entry(frame0, font=("맑은 고딕", 10))
        self.template_entry.insert(0, self.data.get("템플릿_경로", DEFAULT_DATA["템플릿_경로"]))
        self.template_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.entries["템플릿_경로"] = self.template_entry
        
        def browse_template():
            init_dir = os.path.dirname(self.template_entry.get())
            if not os.path.exists(init_dir):
                init_dir = os.path.expanduser("~")
                
            filepath = filedialog.askopenfilename(
                title="원본 한글 양식 선택",
                filetypes=[("한글 파일", "*.hwp"), ("모든 파일", "*.*")],
                initialdir=init_dir
            )
            if filepath:
                self.template_entry.delete(0, tk.END)
                self.template_entry.insert(0, filepath)
                
        btn_browse = ttk.Button(frame0, text="찾아보기...", command=browse_template)
        btn_browse.pack(side="left")

        # 노트북 (탭 컨테이너) 생성 (가운데 확장 영역)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=10)
        
        tab1_scroll = ttk.Frame(self.notebook)
        tab2_scroll = ttk.Frame(self.notebook)
        tab3_scroll = ttk.Frame(self.notebook)
        tab4_scroll = ttk.Frame(self.notebook)
        tab5_scroll = ttk.Frame(self.notebook)
        
        self.notebook.add(tab1_scroll, text="1. 기본 정보")
        self.notebook.add(tab2_scroll, text="2. 위험성평가 현황")
        self.notebook.add(tab3_scroll, text="3. 아차사고 보고서")
        self.notebook.add(tab4_scroll, text="4. 건의 및 제의사항")
        self.notebook.add(tab5_scroll, text="5. 안전보건교육 이수 관리")

        # 필드 생성 도우미 함수
        def create_entry(parent, key, label_text, is_text=False):
            frame = ttk.Frame(parent)
            frame.pack(fill="x", padx=20, pady=5)
            
            lbl = ttk.Label(frame, text=label_text, width=20, anchor="e")
            lbl.pack(side="left", padx=(0, 10))
            
            if is_text:
                widget = tk.Text(frame, height=3, width=40, font=("맑은 고딕", 10))
                widget.insert("1.0", self.data.get(key, ""))
                widget.pack(side="left", fill="x", expand=True)
            else:
                widget = ttk.Entry(frame, width=40, font=("맑은 고딕", 10))
                widget.insert(0, self.data.get(key, ""))
                widget.pack(side="left", fill="x", expand=True)
                
            self.entries[key] = widget

        def create_half_entries(parent, key1, key2, label_text):
            frame = ttk.Frame(parent)
            frame.pack(fill="x", padx=20, pady=5)
            
            lbl = ttk.Label(frame, text=label_text, width=20, anchor="e")
            lbl.pack(side="left", padx=(0, 10))
            
            lbl1 = ttk.Label(frame, text="실시여부:", anchor="e")
            lbl1.pack(side="left", padx=(0, 5))
            
            widget1 = ttk.Entry(frame, width=10, font=("맑은 고딕", 10))
            widget1.insert(0, self.data.get(key1, ""))
            widget1.pack(side="left", padx=(0, 15))
            
            lbl2 = ttk.Label(frame, text="작성날짜:", anchor="e")
            lbl2.pack(side="left", padx=(0, 5))
            
            widget2 = ttk.Entry(frame, width=15, font=("맑은 고딕", 10))
            widget2.insert(0, self.data.get(key2, ""))
            widget2.pack(side="left", fill="x", expand=True)
            
            self.entries[key1] = widget1
            self.entries[key2] = widget2

        # 1. 제출 정보
        lbl_frame1 = ttk.LabelFrame(tab1_scroll, text="제출 정보 및 날짜")
        lbl_frame1.pack(fill="x", padx=20, pady=10)
        create_entry(lbl_frame1, "제출년월", "문서 제출 년/월:")
        create_entry(lbl_frame1, "실적년월", "실적 년/월 (엑셀용):")
        create_entry(lbl_frame1, "계획년월", "계획 년/월 (엑셀용):")

        # 2. 수급업체 현황
        lbl_frame2 = ttk.LabelFrame(tab1_scroll, text="수급업체 현황")
        lbl_frame2.pack(fill="x", padx=20, pady=10)
        create_entry(lbl_frame2, "계약명", "계약 명:")
        create_entry(lbl_frame2, "계약기간", "계약기간:")
        create_entry(lbl_frame2, "계약상대자(업체명)", "업체명:")
        create_entry(lbl_frame2, "현장대리인", "현장대리인:")
        create_entry(lbl_frame2, "작업의 시작시간", "작업 시작시간:")
        create_entry(lbl_frame2, "작업 또는 작업장 간의 연락방법", "연락방법:")
        create_entry(lbl_frame2, "재해발생 위험시의 대피방법", "대피방법:")
        create_entry(lbl_frame2, "사업자와 수급인 또는 수급인 상호간의 연락방법", "상호 연락방법:")

        # 3. 주요 활동 및 요청사항
        lbl_frame3 = ttk.LabelFrame(tab1_scroll, text="주요 활동 및 요청사항")
        lbl_frame3.pack(fill="x", padx=20, pady=10)
        create_entry(lbl_frame3, "작업공정의 조정 및 협의 요청사항", "조정/협의 요청사항:", is_text=True)
        create_entry(lbl_frame3, "주요 활동실적", "이번달 주요 활동실적:", is_text=True)
        create_entry(lbl_frame3, "주요 활동계획", "다음달 주요 활동계획:", is_text=True)
        
        # 4. 위험성평가 현황
        lbl_frame4 = ttk.LabelFrame(tab2_scroll, text="위험성평가 실시 현황")
        lbl_frame4.pack(fill="x", padx=20, pady=10)
        create_half_entries(lbl_frame4, "최초위험성평가_실시여부", "최초위험성평가_작성날짜", "최초 위험성평가:")
        create_half_entries(lbl_frame4, "정기위험성평가_실시여부", "정기위험성평가_작성날짜", "정기 위험성평가:")
        create_half_entries(lbl_frame4, "수시위험성평가_실시여부", "수시위험성평가_작성날짜", "수시 위험성평가:")

        # 5. 추가 위험성평가 내용
        lbl_frame5 = ttk.LabelFrame(tab2_scroll, text="위험성평가 중점관리항목 개선사항")
        lbl_frame5.pack(fill="x", padx=20, pady=10)
        
        create_entry(lbl_frame5, "관리주관부서", "관리 주관부서:")
        create_entry(lbl_frame5, "장소", "장소:")
        create_entry(lbl_frame5, "위험성평가_중점관리항목", "중점관리항목 내용:", is_text=True)
        
        # 사진 1 (조치사진)
        img_frame1 = ttk.Frame(lbl_frame5)
        img_frame1.pack(fill="x", padx=20, pady=5)
        
        img_lbl1 = ttk.Label(img_frame1, text="조치 사진:", width=20, anchor="e")
        img_lbl1.pack(side="left", padx=(0, 10))
        
        self.img_entry1 = ttk.Entry(img_frame1, font=("맑은 고딕", 10))
        self.img_entry1.insert(0, self.data.get("조치사진_경로", ""))
        self.img_entry1.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.entries["조치사진_경로"] = self.img_entry1
        
        def browse_image1():
            init_dir = os.path.dirname(self.img_entry1.get())
            if not init_dir or not os.path.exists(init_dir):
                init_dir = os.path.expanduser("~")
                
            filepaths = filedialog.askopenfilenames(
                title="사진 선택",
                filetypes=[("이미지 파일", "*.jpg *.jpeg *.png *.bmp"), ("모든 파일", "*.*")],
                initialdir=init_dir
            )
            if filepaths:
                self.img_entry1.delete(0, tk.END)
                self.img_entry1.insert(0, "|".join(filepaths))
                
        btn_img_browse1 = ttk.Button(img_frame1, text="사진 찾기...", command=browse_image1)
        btn_img_browse1.pack(side="left")

        # 사진 2 (개선후사진)
        img_frame2 = ttk.Frame(lbl_frame5)
        img_frame2.pack(fill="x", padx=20, pady=5)
        
        img_lbl2 = ttk.Label(img_frame2, text="개선 후 사진:", width=20, anchor="e")
        img_lbl2.pack(side="left", padx=(0, 10))
        
        self.img_entry2 = ttk.Entry(img_frame2, font=("맑은 고딕", 10))
        self.img_entry2.insert(0, self.data.get("개선후사진_경로", ""))
        self.img_entry2.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.entries["개선후사진_경로"] = self.img_entry2
        
        def browse_image2():
            init_dir = os.path.dirname(self.img_entry2.get())
            if not init_dir or not os.path.exists(init_dir):
                init_dir = os.path.expanduser("~")
                
            filepaths = filedialog.askopenfilenames(
                title="사진 선택",
                filetypes=[("이미지 파일", "*.jpg *.jpeg *.png *.bmp"), ("모든 파일", "*.*")],
                initialdir=init_dir
            )
            if filepaths:
                self.img_entry2.delete(0, tk.END)
                self.img_entry2.insert(0, "|".join(filepaths))
                
        btn_img_browse2 = ttk.Button(img_frame2, text="사진 찾기...", command=browse_image2)
        btn_img_browse2.pack(side="left")

        # 6. 위험성 감소대책 Check
        lbl_frame6 = ttk.LabelFrame(tab2_scroll, text="위험성 감소대책 Check")
        lbl_frame6.pack(fill="x", padx=20, pady=10)
        
        check_frame = ttk.Frame(lbl_frame6)
        check_frame.pack(fill="x", padx=20, pady=5)
        
        def create_check(parent, key, text):
            var = tk.BooleanVar(value=self.data.get(key, False))
            cb = ttk.Checkbutton(parent, text=text, variable=var)
            cb.pack(side="left", padx=10)
            self.entries[key] = var
            
        create_check(check_frame, "감소대책_위험성제거", "1. 위험성제거")
        create_check(check_frame, "감소대책_공학적", "2. 공학적/시설적")
        create_check(check_frame, "감소대책_관리적", "3. 관리적 대책")
        create_check(check_frame, "감소대책_개인보호구", "4. 개인보호구")

        # 7. 안전·보건 개선조치 이행사항
        lbl_frame7 = ttk.LabelFrame(tab2_scroll, text="위험성평가에 따른 안전·보건 개선조치 이행사항")
        lbl_frame7.pack(fill="x", padx=20, pady=10)
        create_entry(lbl_frame7, "개선조치_이행사항", "이행사항 내용:", is_text=True)

        # 8. 아차사고 보고서
        lbl_frame8 = ttk.LabelFrame(tab3_scroll, text="아차사고 보고서")
        lbl_frame8.pack(fill="x", padx=20, pady=10)
        
        create_entry(lbl_frame8, "아차사고_사고명", "사고명:")
        create_entry(lbl_frame8, "아차사고_발생일시", "발생일시:")
        create_entry(lbl_frame8, "아차사고_장소", "장소(설비):")
        create_entry(lbl_frame8, "아차사고_보고자", "보고자:")
        create_entry(lbl_frame8, "아차사고_소속", "소속:")
        create_entry(lbl_frame8, "아차사고_사고내용", "사고내용(6하원칙):", is_text=True)
        create_entry(lbl_frame8, "아차사고_원인분석", "원인분석 및 조치의견:", is_text=True)
        
        # 아차사고 사진 1 (조치전)
        img_frame_acha1 = ttk.Frame(lbl_frame8)
        img_frame_acha1.pack(fill="x", padx=20, pady=5)
        
        lbl_acha1 = ttk.Label(img_frame_acha1, text="조치 전 사진:", width=20, anchor="e")
        lbl_acha1.pack(side="left", padx=(0, 10))
        
        self.img_entry_acha1 = ttk.Entry(img_frame_acha1, font=("맑은 고딕", 10))
        self.img_entry_acha1.insert(0, self.data.get("아차사고_조치전사진", ""))
        self.img_entry_acha1.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.entries["아차사고_조치전사진"] = self.img_entry_acha1
        
        def browse_acha1():
            init_dir = os.path.dirname(self.img_entry_acha1.get())
            if not init_dir or not os.path.exists(init_dir):
                init_dir = os.path.expanduser("~")
            filepaths = filedialog.askopenfilenames(title="아차사고 조치 전 사진 선택", filetypes=[("이미지 파일", "*.jpg *.jpeg *.png *.bmp"), ("모든 파일", "*.*")], initialdir=init_dir)
            if filepaths:
                self.img_entry_acha1.delete(0, tk.END)
                self.img_entry_acha1.insert(0, "|".join(filepaths))
                
        btn_acha1 = ttk.Button(img_frame_acha1, text="사진 찾기...", command=browse_acha1)
        btn_acha1.pack(side="left")

        # 아차사고 사진 2 (조치후)
        img_frame_acha2 = ttk.Frame(lbl_frame8)
        img_frame_acha2.pack(fill="x", padx=20, pady=5)
        
        lbl_acha2 = ttk.Label(img_frame_acha2, text="조치 후 사진:", width=20, anchor="e")
        lbl_acha2.pack(side="left", padx=(0, 10))
        
        self.img_entry_acha2 = ttk.Entry(img_frame_acha2, font=("맑은 고딕", 10))
        self.img_entry_acha2.insert(0, self.data.get("아차사고_조치후사진", ""))
        self.img_entry_acha2.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.entries["아차사고_조치후사진"] = self.img_entry_acha2
        
        def browse_acha2():
            init_dir = os.path.dirname(self.img_entry_acha2.get())
            if not init_dir or not os.path.exists(init_dir):
                init_dir = os.path.expanduser("~")
            filepaths = filedialog.askopenfilenames(title="아차사고 조치 후 사진 선택", filetypes=[("이미지 파일", "*.jpg *.jpeg *.png *.bmp"), ("모든 파일", "*.*")], initialdir=init_dir)
            if filepaths:
                self.img_entry_acha2.delete(0, tk.END)
                self.img_entry_acha2.insert(0, "|".join(filepaths))
                
        btn_acha2 = ttk.Button(img_frame_acha2, text="사진 찾기...", command=browse_acha2)
        btn_acha2.pack(side="left")

        # 9. 건의 및 제의사항
        lbl_frame9 = ttk.LabelFrame(tab4_scroll, text="안전·보건관련 건의 및 제의사항")
        lbl_frame9.pack(fill="x", padx=20, pady=10)
        
        create_entry(lbl_frame9, "건의_개진사항", "개진사항:", is_text=True)
        create_entry(lbl_frame9, "건의_제안사유", "제안사유:", is_text=True)

        # ── 탭5: 안전보건교육 이수 관리대장 ──────────────────────────────
        lbl_frame_edu_header = ttk.LabelFrame(tab5_scroll, text="관리대장 기본 정보")
        lbl_frame_edu_header.pack(fill="x", padx=20, pady=10)
        create_entry(lbl_frame_edu_header, "교육관리_년월",    "교육 년/월:")
        create_entry(lbl_frame_edu_header, "교육관리_수급업체명", "수급업체명:")
        create_entry(lbl_frame_edu_header, "교육관리_계약명",  "계약명:")
        create_entry(lbl_frame_edu_header, "교육관리_현장대리인", "현장대리인:")

        # 직원 목록 Treeview
        lbl_frame_edu = ttk.LabelFrame(tab5_scroll, text="직원 교육 이수 현황")
        lbl_frame_edu.pack(fill="both", expand=True, padx=20, pady=5)

        # Treeview 컬럼 정의
        edu_cols = ("no", "name", "birth", "type", "hired",
                    "채용시교육", "작업내용변경시", "특별교육",
                    "q1_1", "q1_2", "q1_3",
                    "q2_4", "q2_5", "q2_6",
                    "q3_7", "q3_8", "q3_9",
                    "q4_10", "q4_11", "q4_12")
        edu_col_labels = {
            "no": "NO", "name": "성명", "birth": "생년월일",
            "type": "구분\n(관리/근로)", "hired": "채용일자(최초입일)",
            "채용시교육": "채용시\n(1시간)", "작업내용변경시": "작업내용\n변경시(2h)", "특별교육": "특별교육\n(16시간이상)",
            "q1_1": "1월", "q1_2": "2월", "q1_3": "3월",
            "q2_4": "4월", "q2_5": "5월", "q2_6": "6월",
            "q3_7": "7월", "q3_8": "8월", "q3_9": "9월",
            "q4_10": "10월", "q4_11": "11월", "q4_12": "12월"
        }

        tree_frame = ttk.Frame(lbl_frame_edu)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal")
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical")

        self.edu_tree = ttk.Treeview(
            tree_frame,
            columns=edu_cols,
            show="headings",
            height=10,
            xscrollcommand=xscroll.set,
            yscrollcommand=yscroll.set
        )
        xscroll.config(command=self.edu_tree.xview)
        yscroll.config(command=self.edu_tree.yview)

        col_widths = {
            "no": 35, "name": 65, "birth": 70, "type": 75, "hired": 80,
            "채용시교육": 60, "작업내용변경시": 75, "특별교육": 75,
            "q1_1": 40, "q1_2": 40, "q1_3": 40,
            "q2_4": 40, "q2_5": 40, "q2_6": 40,
            "q3_7": 40, "q3_8": 40, "q3_9": 40,
            "q4_10": 40, "q4_11": 40, "q4_12": 40
        }
        for col in edu_cols:
            self.edu_tree.heading(col, text=edu_col_labels[col])
            self.edu_tree.column(col, width=col_widths[col], anchor="center", minwidth=col_widths[col])

        self.edu_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # 기존 데이터 로드
        def refresh_edu_tree():
            for item in self.edu_tree.get_children():
                self.edu_tree.delete(item)
            for i, r in enumerate(self.edu_records, 1):
                self.edu_tree.insert("", "end", iid=str(i), values=(
                    i,
                    r.get("name", ""), r.get("birth", ""),
                    r.get("type", ""), r.get("hired", ""),
                    r.get("채용시교육", ""), r.get("작업내용변경시", ""), r.get("특별교육", ""),
                    r.get("q1_1", ""), r.get("q1_2", ""), r.get("q1_3", ""),
                    r.get("q2_4", ""), r.get("q2_5", ""), r.get("q2_6", ""),
                    r.get("q3_7", ""), r.get("q3_8", ""), r.get("q3_9", ""),
                    r.get("q4_10", ""), r.get("q4_11", ""), r.get("q4_12", "")
                ))
        refresh_edu_tree()

        # 입력/수정 다이얼로그
        def open_edu_dialog(record=None, iid=None):
            dlg = tk.Toplevel(self.root)
            dlg.title("직원 교육 이수 정보 입력")
            dlg.grab_set()
            dlg.resizable(False, False)

            fields_basic = [
                ("name", "성명"),
                ("birth", "생년월일 (예: 77.07.07)"),
                ("type", "구분 (관리감독자 / 근로자)"),
                ("hired", "채용일자 (예: 20-01-01)"),
                ("채용시교육", "채용시 교육 이수 (이수/미이수)"),
                ("작업내용변경시", "작업내용변경시 교육 이수 (이수/미이수)"),
                ("특별교육", "특별교육 이수 (이수/미이수)"),
            ]
            month_fields = [
                ("q1_1","1월"),("q1_2","2월"),("q1_3","3월"),
                ("q2_4","4월"),("q2_5","5월"),("q2_6","6월"),
                ("q3_7","7월"),("q3_8","8월"),("q3_9","9월"),
                ("q4_10","10월"),("q4_11","11월"),("q4_12","12월"),
            ]

            dlg_entries = {}
            frm = ttk.Frame(dlg, padding=15)
            frm.pack(fill="both", expand=True)

            ttk.Label(frm, text="■ 기본 정보", font=("맑은 고딕", 10, "bold")).grid(
                row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

            for r_idx, (key, label) in enumerate(fields_basic, 1):
                ttk.Label(frm, text=label + ":", anchor="e", width=28).grid(
                    row=r_idx, column=0, sticky="e", padx=(0, 8), pady=3)
                e = ttk.Entry(frm, width=30, font=("맑은 고딕", 10))
                e.grid(row=r_idx, column=1, sticky="ew", pady=3)
                if record:
                    e.insert(0, record.get(key, ""))
                dlg_entries[key] = e

            sep_row = len(fields_basic) + 1
            ttk.Separator(frm, orient="horizontal").grid(
                row=sep_row, column=0, columnspan=2, sticky="ew", pady=8)
            ttk.Label(frm, text="■ 월별 정기교육 이수 시간 (시간 단위로 입력)",
                      font=("맑은 고딕", 10, "bold")).grid(
                row=sep_row + 1, column=0, columnspan=2, sticky="w", pady=(0, 5))

            month_frame = ttk.Frame(frm)
            month_frame.grid(row=sep_row + 2, column=0, columnspan=2, sticky="ew")
            for m_idx, (key, label) in enumerate(month_fields):
                ttk.Label(month_frame, text=label + ":", width=5, anchor="e").grid(
                    row=m_idx // 4, column=(m_idx % 4) * 2, sticky="e", padx=(4, 2), pady=3)
                e = ttk.Entry(month_frame, width=6, font=("맑은 고딕", 10))
                e.grid(row=m_idx // 4, column=(m_idx % 4) * 2 + 1, sticky="w", padx=(0, 8), pady=3)
                if record:
                    e.insert(0, record.get(key, ""))
                dlg_entries[key] = e

            def on_ok():
                new_rec = {k: v.get().strip() for k, v in dlg_entries.items()}
                if not new_rec["name"]:
                    messagebox.showwarning("입력 오류", "성명은 필수 입력입니다.", parent=dlg)
                    return
                if iid is not None:
                    self.edu_records[iid] = new_rec
                else:
                    self.edu_records.append(new_rec)
                refresh_edu_tree()
                dlg.destroy()

            btn_ok = ttk.Button(frm, text="확인", command=on_ok)
            btn_ok.grid(row=sep_row + 3, column=0, columnspan=2, pady=12, ipadx=20)
            frm.columnconfigure(1, weight=1)

        def add_edu_record():
            open_edu_dialog()

        def edit_edu_record():
            sel = self.edu_tree.selection()
            if not sel:
                messagebox.showinfo("알림", "수정할 직원을 선택하세요.")
                return
            idx = int(sel[0]) - 1
            open_edu_dialog(record=self.edu_records[idx], iid=idx)

        def delete_edu_record():
            sel = self.edu_tree.selection()
            if not sel:
                messagebox.showinfo("알림", "삭제할 직원을 선택하세요.")
                return
            if messagebox.askyesno("삭제 확인", "선택한 직원 기록을 삭제하시겠습니까?"):
                idx = int(sel[0]) - 1
                self.edu_records.pop(idx)
                refresh_edu_tree()

        edu_btn_frame = ttk.Frame(lbl_frame_edu)
        edu_btn_frame.pack(fill="x", padx=5, pady=(0, 5))
        ttk.Button(edu_btn_frame, text="➕ 직원 추가", command=add_edu_record).pack(side="left", padx=3)
        ttk.Button(edu_btn_frame, text="✏️ 수정",     command=edit_edu_record).pack(side="left", padx=3)
        ttk.Button(edu_btn_frame, text="🗑️ 삭제",    command=delete_edu_record).pack(side="left", padx=3)

        # 버튼 프레임 (최하단 고정)
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=20, pady=10)

        excel_btn = ttk.Button(btn_frame, text="엑셀(.xlsx)로 완벽하게 생성 (추천!)", command=self.do_generate_excel)
        excel_btn.pack(side="left", fill="x", expand=True, padx=5, ipady=10)

        hwp_btn = ttk.Button(btn_frame, text="한글(.hwp) 파일 생성", command=self.do_generate_hwp)
        hwp_btn.pack(side="left", fill="x", expand=True, padx=5, ipady=10)

    def do_generate_excel(self):
        self.save_config()
        data = self.get_current_data()
        desktop_path = os.path.join(os.path.expanduser("~"), "OneDrive", "바탕 화면")
        base_name = f"{data['제출년월'].replace(' ', '_')}_안전보건협의체_회의자료"

        # 파일이 이미 열려있는 경우 _2, _3 ... 으로 자동 대체
        output_path = os.path.join(desktop_path, base_name + ".xlsx")
        counter = 2
        while True:
            try:
                generate_excel(data, output_path)
                break
            except PermissionError:
                alt = os.path.join(desktop_path, f"{base_name}_{counter}.xlsx")
                if messagebox.askyesno(
                    "파일 사용 중",
                    f"'{os.path.basename(output_path)}' 파일이 열려 있어 저장할 수 없습니다.\n\n"
                    f"Excel에서 해당 파일을 닫거나,\n"
                    f"'{os.path.basename(alt)}' 이름으로 저장할까요?\n\n"
                    f"[예] → 새 이름으로 저장\n[아니오] → 취소"
                ):
                    output_path = alt
                    counter += 1
                else:
                    return
            except Exception as e:
                messagebox.showerror("오류", f"엑셀 생성 중 오류가 발생했습니다:\n{e}")
                return

        messagebox.showinfo(
            "성공",
            f"엑셀 파일이 바탕화면에 생성되었습니다.\n\n"
            f"경로: {output_path}\n\n"
            f"시트1: 안전보건협의체 회의자료\n"
            f"시트2: 안전보건교육 이수 관리대장"
        )

    def do_generate_hwp(self):
        # 저장
        self.save_config()
        data = self.get_current_data()
        out_name = f"{data['제출년월'].replace(' ', '_')}_안전보건협의체_회의자료_작성본.hwp"
        desktop_path = os.path.join(os.path.expanduser("~"), "OneDrive", "바탕 화면")
        output_path = os.path.join(desktop_path, out_name)
        
        template_hwp = self.entries["템플릿_경로"].get()
        if not os.path.exists(template_hwp):
            messagebox.showerror("오류", f"한글 템플릿 파일이 존재하지 않습니다.\n{template_hwp}")
            return
            
        try:
            generate_hwp(data, template_hwp, output_path)
            msg = (f"한글 파일이 바탕화면에 생성되었습니다.\n\n경로: {output_path}\n\n"
                   f"주의: 한글 양식 내부에 누름틀(필드 이름)이 없는 경우 "
                   f"'주요 활동실적' 등 줄바꿈이 있는 항목은 좌측 셀에 입력될 수 있습니다.")
            messagebox.showinfo("성공", msg)
        except Exception as e:
            messagebox.showerror("오류", f"한글 생성 중 오류가 발생했습니다:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    
    # 테마 설정 (옵션)
    style = ttk.Style(root)
    style.theme_use('clam')
    
    app = FormGeneratorApp(root)
    root.mainloop()
