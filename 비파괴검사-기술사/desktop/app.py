from __future__ import annotations

import json
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz
from PIL import Image, ImageTk

from study_content import WEEKLY_CONTENT


APP_DIR = Path(__file__).resolve().parent
DATA_FILE = APP_DIR / "study_data.json"
SETTINGS_FILE = APP_DIR / "settings.json"
PROBLEM_INDEX_FILE = APP_DIR / "problem_index.json"
DEFAULT_SOURCE_DIR = Path(r"I:\주진철\도서\자격증\비파괴검사 기술사")
SOURCE_FILES = ["기술사 2022-1.pdf", "기술사 2022-2.pdf", "기술사 2022-3.pdf"]

COLORS = {
    "ink": "#10231d",
    "green": "#0b5d48",
    "green2": "#14785e",
    "lime": "#c8e56a",
    "paper": "#f4f2eb",
    "white": "#ffffff",
    "line": "#deded5",
    "muted": "#68766f",
    "soft": "#e7ebdd",
    "amber": "#e9a23b",
    "red": "#c95d51",
}

WEEKS = [
    ("초음파 탐상검사 UT", "파동, 음향임피던스, 근거리음장, 감쇠, 탐촉자"),
    ("방사선투과 RT", "X선·감마선, 투과도계, 필름, CR·DR, 방사선 방호"),
    ("자분·침투 MT/PT", "자화법, 표준시험편, 침투제, 유화제, 현상제"),
    ("와전류·누설 ET/LT", "코일, 침투깊이, 리프트오프, 헬륨, 진공상자법"),
    ("육안·음향방출 VT/AE", "조명, 시력, AE 신호, 카이저·펠리시티 효과"),
    ("용접", "결함, 저온·고온균열, 잔류응력, PWHT, 입열"),
    ("금속재료", "열처리, HAZ, 취성, 피로, 부식, 스테인리스강"),
    ("고급 초음파", "PAUT, TOFD, Guided Wave, EMAT, 레이저 UT"),
    ("설비 적용", "배관, 보일러, 저장탱크, 복수기, 원전설비"),
    ("건전성 평가", "POD, 기량검증, 파괴역학, RBI, 잔여수명"),
    ("법규·규격", "원자력안전법, 방사선량, ISO 9712, KS, ASME"),
    ("실전 종합", "100분 답안 훈련, 약점 보완, 목차 암기"),
]

QUESTIONS = [
    ("q-paut", "UT", "위상배열 초음파탐상검사(PAUT)의 원리와 특성을 설명하시오."),
    ("q-stb", "UT", "STB-A1 표준시험편의 용도를 설명하시오."),
    ("q-near", "UT", "근거리음장 한계거리와 탐상에 미치는 영향을 설명하시오."),
    ("q-imp", "UT", "음향임피던스와 경계면에서의 반사·투과 관계를 설명하시오."),
    ("q-tofd", "UT", "TOFD의 원리와 결함 크기 측정방법을 설명하시오."),
    ("q-couplant", "UT", "접촉매질의 사용목적과 선정 시 고려인자를 설명하시오."),
    ("q-ray", "RT", "알파·베타·감마선의 특성과 물질과의 상호작용을 비교하시오."),
    ("q-iqi", "RT", "방사선투과검사에서 투과도계를 사용하는 목적을 설명하시오."),
    ("q-cr", "RT", "CR의 원리와 특징을 설명하시오."),
    ("q-alara", "RT", "방사선 방호 3원칙과 ALARA 개념을 설명하시오."),
    ("q-mt", "MT", "자분탐상검사의 원리와 장단점을 설명하시오."),
    ("q-mtblock", "MT", "A형 표준시험편과 B형 대비시험편의 용도를 비교하시오."),
    ("q-skin", "MT", "교류자화에서의 표피효과를 설명하시오."),
    ("q-pt", "PT", "침투탐상검사의 원리와 장단점을 설명하시오."),
    ("q-wet", "PT", "침투액의 적심성과 모세관현상을 설명하시오."),
    ("q-pttemp", "PT", "표준온도를 벗어난 침투탐상검사의 인정요건을 설명하시오."),
    ("q-et", "ET", "와전류탐상검사의 원리와 장단점을 설명하시오."),
    ("q-liftoff", "ET", "와전류검사에서 리프트오프 효과를 설명하시오."),
    ("q-rfect", "ET", "원격장 와전류탐상(RFECT)의 원리와 적용을 설명하시오."),
    ("q-ae", "AE/VT/LT", "AE와 UT를 에너지원·신호·결함평가 관점에서 비교하시오."),
    ("q-kaiser", "AE/VT/LT", "카이저 효과와 펠리시티 효과를 설명하시오."),
    ("q-helium", "AE/VT/LT", "헬륨누설시험의 원리와 특징을 설명하시오."),
    ("q-vt", "AE/VT/LT", "직접·간접·투시 육안검사의 적용기법을 설명하시오."),
    ("q-ce", "용접·재료", "탄소당량의 산출식과 활용을 설명하시오."),
    ("q-crack", "용접·재료", "용접부 저온균열의 특징과 방지대책을 설명하시오."),
    ("q-pwht", "용접·재료", "용접후열처리(PWHT)의 목적을 설명하시오."),
    ("q-sens", "용접·재료", "스테인리스강의 예민화와 입계부식 방지대책을 설명하시오."),
    ("q-pod", "건전성·법규", "POD 곡선의 의미와 결정방법을 설명하시오."),
    ("q-rbi", "건전성·법규", "위험기반검사(RBI)의 개념과 절차를 설명하시오."),
    ("q-iso", "건전성·법규", "ISO 9712 Level 3 자격인정 체계를 설명하시오."),
]


def load_problem_index() -> list[dict]:
    try:
        problems = json.loads(PROBLEM_INDEX_FILE.read_text(encoding="utf-8"))
        return sorted(problems, key=lambda p: (p.get("week", 12), p.get("study_group", ""), p["number"]))
    except (OSError, json.JSONDecodeError):
        return []


ALL_PROBLEMS = load_problem_index()


def load_data() -> dict:
    empty = {"weeks": [], "questions": [], "tasks": [], "answers": {}, "mistakes": []}
    if not DATA_FILE.exists():
        return empty
    try:
        saved = json.loads(DATA_FILE.read_text(encoding="utf-8"))
        empty.update(saved)
    except (OSError, json.JSONDecodeError):
        pass
    return empty


def load_settings() -> dict:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


class ScrollFrame(tk.Frame):
    def __init__(self, parent, bg: str = COLORS["paper"]):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.content = tk.Frame(self.canvas, bg=bg)
        self.window = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.content.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.window, width=e.width))
        self.canvas.bind_all("<MouseWheel>", self._wheel)

    def _wheel(self, event):
        if self.winfo_ismapped():
            self.canvas.yview_scroll(int(-event.delta / 120), "units")


class StudyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NDT 기술사 학습실")
        self.geometry("1280x820")
        self.minsize(1040, 680)
        self.configure(bg=COLORS["paper"])
        self.data = load_data()
        settings = load_settings()
        portable_books = APP_DIR / "교재"
        configured = Path(settings["source_dir"]) if settings.get("source_dir") else None
        if configured and configured.exists():
            self.source_dir = configured
        elif portable_books.exists():
            self.source_dir = portable_books
        else:
            self.source_dir = DEFAULT_SOURCE_DIR
        self.frames: dict[str, tk.Widget] = {}
        self.nav_buttons: dict[str, tk.Button] = {}
        self.timer_seconds = 1500
        self.timer_running = False
        self.timer_job = None
        self._configure_styles()
        self._build_shell()
        self.show_view("dashboard")
        self.protocol("WM_DELETE_WINDOW", self.close_app)

    def _configure_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TCombobox", padding=8, fieldbackground=COLORS["white"])
        style.map("TCombobox", fieldbackground=[("readonly", COLORS["white"])])

    def _build_shell(self):
        sidebar = tk.Frame(self, bg=COLORS["ink"], width=238)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        brand = tk.Frame(sidebar, bg=COLORS["ink"])
        brand.pack(fill="x", padx=24, pady=(30, 42))
        tk.Label(brand, text="N", bg=COLORS["ink"], fg=COLORS["lime"], font=("Georgia", 24)).pack(side="left", padx=(0, 12))
        brand_text = tk.Frame(brand, bg=COLORS["ink"])
        brand_text.pack(side="left")
        tk.Label(brand_text, text="NDT 기술사", bg=COLORS["ink"], fg="white", font=("맑은 고딕", 14, "bold")).pack(anchor="w")
        tk.Label(brand_text, text="STUDY WORKSPACE", bg=COLORS["ink"], fg="#91a49c", font=("Arial", 8)).pack(anchor="w")

        items = [("dashboard", "01   대시보드"), ("plan", "02   12주 학습계획"), ("theory", "03   주차별 학습답안"), ("questions", "04   문제은행"), ("answer", "05   답안연습"), ("mistakes", "06   오답노트")]
        for key, label in items:
            button = tk.Button(sidebar, text=label, anchor="w", padx=24, pady=12, bd=0, bg=COLORS["ink"], fg="#b7c3be", activebackground="#1d3931", activeforeground="white", font=("맑은 고딕", 10), command=lambda k=key: self.show_view(k))
            button.pack(fill="x", padx=12, pady=2)
            self.nav_buttons[key] = button

        source = tk.Frame(sidebar, bg="#17372e", padx=16, pady=15)
        source.pack(side="bottom", fill="x", padx=18, pady=22)
        tk.Label(source, text="현재 교재", bg="#17372e", fg=COLORS["lime"], font=("맑은 고딕", 8, "bold")).pack(anchor="w")
        tk.Label(source, text="PERFECT 2022 (1)-(3)", bg="#17372e", fg="white", font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(5, 0))
        tk.Label(source, text=f"전체문제 {len(ALL_PROBLEMS)}개", bg="#17372e", fg="#91a49c", font=("맑은 고딕", 8)).pack(anchor="w")
        self.source_status = tk.Label(source, text="", bg="#17372e", fg="#91a49c", font=("맑은 고딕", 8), wraplength=180, justify="left")
        self.source_status.pack(anchor="w", pady=(5, 8))
        tk.Button(source, text="교재 폴더 설정", command=self.choose_source_dir, bd=0, bg=COLORS["lime"], fg=COLORS["ink"], activebackground="#b7d45c", padx=10, pady=6, font=("맑은 고딕", 8, "bold")).pack(anchor="w")
        self.update_source_status()

        body = tk.Frame(self, bg=COLORS["paper"])
        body.pack(side="left", fill="both", expand=True)
        header = tk.Frame(body, bg=COLORS["paper"], height=98)
        header.pack(fill="x", padx=38)
        header.pack_propagate(False)
        title_box = tk.Frame(header, bg=COLORS["paper"])
        title_box.pack(side="left", pady=22)
        tk.Label(title_box, text="비파괴검사 기술사", bg=COLORS["paper"], fg=COLORS["green2"], font=("맑은 고딕", 8, "bold")).pack(anchor="w")
        self.page_title = tk.Label(title_box, text="", bg=COLORS["paper"], fg=COLORS["ink"], font=("맑은 고딕", 19, "bold"))
        self.page_title.pack(anchor="w")
        tk.Label(header, text=date.today().strftime("%Y.%m.%d"), bg=COLORS["paper"], fg=COLORS["muted"], font=("Arial", 10)).pack(side="right")
        self.view_host = tk.Frame(body, bg=COLORS["paper"])
        self.view_host.pack(fill="both", expand=True, padx=38, pady=(0, 32))

    def save(self):
        DATA_FILE.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")

    def close_app(self):
        self.save()
        self.destroy()

    def source_files_ready(self, directory=None):
        folder = Path(directory) if directory else self.source_dir
        return all((folder / name).exists() for name in SOURCE_FILES)

    def update_source_status(self):
        if self.source_files_ready():
            self.source_status.config(text="원문 PDF 연결됨", fg=COLORS["lime"])
        else:
            self.source_status.config(text="원문 PDF 미연결", fg="#e9a23b")

    def choose_source_dir(self):
        selected = filedialog.askdirectory(title="기술사 PDF 3개가 있는 폴더 선택", initialdir=str(self.source_dir) if self.source_dir.exists() else str(APP_DIR))
        if not selected:
            return
        folder = Path(selected)
        missing = [name for name in SOURCE_FILES if not (folder / name).exists()]
        if missing:
            messagebox.showwarning("교재 파일 확인", "선택한 폴더에 다음 파일이 없습니다.\n\n" + "\n".join(missing))
            return
        self.source_dir = folder
        SETTINGS_FILE.write_text(json.dumps({"source_dir": str(folder)}, ensure_ascii=False, indent=2), encoding="utf-8")
        self.update_source_status()
        messagebox.showinfo("교재 연결 완료", "원문 PDF 폴더를 저장했습니다.")

    def clear_host(self):
        for child in self.view_host.winfo_children():
            child.destroy()

    def show_view(self, name: str):
        titles = {"dashboard": "학습 대시보드", "plan": "12주 학습계획", "theory": "주차별 학습답안", "questions": "문제은행", "answer": "답안연습", "mistakes": "오답노트"}
        self.page_title.config(text=titles[name])
        for key, button in self.nav_buttons.items():
            button.config(bg="#1d3931" if key == name else COLORS["ink"], fg="white" if key == name else "#b7c3be")
        self.clear_host()
        getattr(self, f"build_{name}")()

    def label(self, parent, text, size=10, color=None, bold=False, bg=None, **kwargs):
        return tk.Label(parent, text=text, bg=bg or parent.cget("bg"), fg=color or COLORS["ink"], font=("맑은 고딕", size, "bold" if bold else "normal"), **kwargs)

    def card(self, parent, bg=None, **pack):
        frame = tk.Frame(parent, bg=bg or COLORS["white"], highlightbackground=COLORS["line"], highlightthickness=1, padx=22, pady=18)
        frame.pack(**pack)
        return frame

    def action_button(self, parent, text, command, primary=True):
        return tk.Button(parent, text=text, command=command, bd=0, padx=16, pady=9, cursor="hand2", bg=COLORS["lime"] if primary else COLORS["white"], fg=COLORS["ink"], activebackground="#b7d45c" if primary else COLORS["soft"], font=("맑은 고딕", 9, "bold"))

    def build_dashboard(self):
        container = ScrollFrame(self.view_host)
        container.pack(fill="both", expand=True)
        root = container.content
        hero = tk.Frame(root, bg=COLORS["green"], padx=38, pady=30)
        hero.pack(fill="x", pady=(0, 16))
        left = tk.Frame(hero, bg=COLORS["green"])
        left.pack(side="left", fill="both", expand=True)
        self.label(left, "12주 집중 과정", 9, COLORS["lime"], True, COLORS["green"]).pack(anchor="w")
        self.label(left, "합격 답안은 반복해서 완성됩니다.", 23, "white", True, COLORS["green"]).pack(anchor="w", pady=(8, 5))
        self.label(left, "오늘의 이론을 익히고, 한 문제를 직접 써보세요.", 10, "#c4d5ce", bg=COLORS["green"]).pack(anchor="w")
        self.action_button(left, "오늘 답안 시작하기  →", lambda: self.open_question("q-paut")).pack(anchor="w", pady=(20, 0))
        total = len(WEEKS) + len(ALL_PROBLEMS) + 4
        done = len(self.data["weeks"]) + len(self.data["questions"]) + len(self.data["tasks"])
        percent = round(done / total * 100)
        progress = tk.Frame(hero, bg="#164f40", padx=26, pady=20)
        progress.pack(side="right", padx=(24, 0))
        self.label(progress, f"{percent}%", 28, "white", True, "#164f40").pack()
        self.label(progress, "전체 진도", 9, "#c4d5ce", bg="#164f40").pack()

        metrics = tk.Frame(root, bg=COLORS["paper"])
        metrics.pack(fill="x", pady=(0, 16))
        values = [("완료 주차", f'{len(self.data["weeks"])} / 12'), ("완료 문제", f'{len(self.data["questions"])} / {len(ALL_PROBLEMS)}'), ("작성 답안", f'{len(self.data["answers"])}개'), ("오답 기록", f'{len(self.data["mistakes"])}개')]
        for i, (title, value) in enumerate(values):
            box = tk.Frame(metrics, bg=COLORS["white"], highlightbackground=COLORS["line"], highlightthickness=1, padx=18, pady=14)
            box.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 5, 0 if i == 3 else 5))
            metrics.grid_columnconfigure(i, weight=1)
            self.label(box, title, 8, COLORS["muted"]).pack(anchor="w")
            self.label(box, value, 18, COLORS["ink"], True).pack(anchor="w", pady=(4, 0))

        tasks = self.card(root, fill="x")
        self.label(tasks, "이번 주 · 초음파탐상 UT", 14, bold=True).pack(anchor="w", pady=(0, 8))
        for task_id, text in [("ut-theory", "파동과 음향임피던스"), ("ut-nearfield", "근거리음장과 감쇠"), ("ut-probe", "탐촉자와 주파수 선정"), ("ut-paut", "PAUT 원리와 특성")]:
            var = tk.BooleanVar(value=task_id in self.data["tasks"])
            cb = tk.Checkbutton(tasks, text=text, variable=var, bg=COLORS["white"], activebackground=COLORS["white"], anchor="w", font=("맑은 고딕", 10), command=lambda t=task_id, v=var: self.toggle_list("tasks", t, v.get(), refresh=False))
            cb.pack(fill="x", pady=5)

    def build_plan(self):
        scroll = ScrollFrame(self.view_host)
        scroll.pack(fill="both", expand=True)
        root = scroll.content
        self.label(root, "완료한 주차를 체크하면 전체 진도에 반영됩니다.", 10, COLORS["muted"]).pack(anchor="w", pady=(0, 14))
        for i, (title, detail) in enumerate(WEEKS, 1):
            card = tk.Frame(root, bg=COLORS["soft"] if i in self.data["weeks"] else COLORS["white"], highlightbackground=COLORS["line"], highlightthickness=1, padx=20, pady=15)
            card.pack(fill="x", pady=5)
            text = tk.Frame(card, bg=card.cget("bg"))
            text.pack(side="left", fill="x", expand=True)
            self.label(text, f"W{i:02d}  {title}", 12, bold=True, bg=card.cget("bg")).pack(anchor="w")
            self.label(text, detail, 9, COLORS["muted"], bg=card.cget("bg")).pack(anchor="w", pady=(4, 0))
            var = tk.BooleanVar(value=i in self.data["weeks"])
            tk.Checkbutton(card, text="완료", variable=var, bg=card.cget("bg"), activebackground=card.cget("bg"), command=lambda n=i, v=var: self.toggle_list("weeks", n, v.get(), refresh=True)).pack(side="right")

    def build_questions(self):
        top = tk.Frame(self.view_host, bg=COLORS["paper"])
        top.pack(fill="x", pady=(0, 12))
        self.search_var = tk.StringVar()
        self.study_week_var = tk.StringVar(value="전체 주차")
        self.study_group_var = tk.StringVar(value="전체 원리")
        self.answer_type_var = tk.StringVar(value="전체 유형")
        search = tk.Entry(top, textvariable=self.search_var, font=("맑은 고딕", 10), relief="solid", bd=1)
        search.pack(side="left", fill="x", expand=True, ipady=8)
        week_values = ["전체 주차"] + [f'{week:02d}주 {name}' for week, name in sorted({(p["week"], p["week_name"]) for p in ALL_PROBLEMS})]
        week_combo = ttk.Combobox(top, textvariable=self.study_week_var, values=week_values, state="readonly", width=22)
        week_combo.pack(side="left", padx=(10, 0))
        filters = tk.Frame(self.view_host, bg=COLORS["paper"])
        filters.pack(fill="x", pady=(0, 10))
        self.group_combo = ttk.Combobox(filters, textvariable=self.study_group_var, state="readonly", width=28)
        self.group_combo.pack(side="left")
        type_values = ["전체 유형"] + sorted({p["answer_type"] for p in ALL_PROBLEMS})
        type_combo = ttk.Combobox(filters, textvariable=self.answer_type_var, values=type_values, state="readonly", width=18)
        type_combo.pack(side="left", padx=(10, 0))
        self.label(filters, "주차 → 핵심원리 → 답안유형 순으로 좁혀서 공부하세요.", 9, COLORS["muted"]).pack(side="left", padx=14)
        self.search_var.trace_add("write", lambda *_: self.render_question_list())
        week_combo.bind("<<ComboboxSelected>>", lambda _e: self.on_week_filter_changed())
        self.group_combo.bind("<<ComboboxSelected>>", lambda _e: self.render_question_list())
        type_combo.bind("<<ComboboxSelected>>", lambda _e: self.render_question_list())
        self.refresh_group_options()
        self.question_scroll = ScrollFrame(self.view_host)
        self.question_scroll.pack(fill="both", expand=True)
        self.render_question_list()

    def selected_study_week(self):
        value = self.study_week_var.get()
        return None if value == "전체 주차" else int(value[:2])

    def refresh_group_options(self):
        week = self.selected_study_week()
        groups = sorted({p["study_group"] for p in ALL_PROBLEMS if week is None or p["week"] == week})
        values = ["전체 원리"] + groups
        self.group_combo.configure(values=values)
        if self.study_group_var.get() not in values:
            self.study_group_var.set("전체 원리")

    def on_week_filter_changed(self):
        self.refresh_group_options()
        self.render_question_list()

    def show_similar_group(self, problem):
        self.study_week_var.set(f'{problem["week"]:02d}주 {problem["week_name"]}')
        self.refresh_group_options()
        self.study_group_var.set(problem["study_group"])
        self.answer_type_var.set("전체 유형")
        self.render_question_list()

    def build_theory(self):
        body = tk.Frame(self.view_host, bg=COLORS["paper"])
        body.pack(fill="both", expand=True)
        selector = tk.Frame(body, bg=COLORS["paper"], width=305)
        selector.pack(side="left", fill="y", padx=(0, 14))
        selector.pack_propagate(False)
        self.theory_selector = selector
        self.theory_body = body
        self.label(selector, "12주 답안 목록", 12, bold=True).pack(anchor="w", pady=(0, 10))
        self.theory_week_list = tk.Listbox(selector, bd=0, highlightthickness=1, highlightbackground=COLORS["line"], selectbackground=COLORS["green"], selectforeground="white", font=("맑은 고딕", 10), activestyle="none")
        self.theory_week_list.pack(fill="both", expand=True)
        for item in WEEKLY_CONTENT:
            self.theory_week_list.insert("end", f'{item["week"]:02d}주  {item["theme"]}')
        self.theory_panel = tk.Frame(body, bg=COLORS["paper"])
        self.theory_panel.pack(side="left", fill="both", expand=True)
        selected = min(max(getattr(self, "pending_theory_week", 1), 1), len(WEEKLY_CONTENT)) - 1
        self.theory_week_list.selection_set(selected)
        self.theory_week_list.see(selected)
        self.render_theory_week()
        self.theory_week_list.bind("<<ListboxSelect>>", self.on_theory_week_selected)

    def on_theory_week_selected(self, _event=None):
        selection = self.theory_week_list.curselection()
        if not selection:
            return
        selected_week = WEEKLY_CONTENT[selection[0]]["week"]
        if selected_week == getattr(self, "pending_theory_week", None):
            return
        self.render_theory_week()

    def render_theory_week(self):
        for child in self.theory_panel.winfo_children():
            child.destroy()
        selection = self.theory_week_list.curselection()
        if not selection:
            return
        item = WEEKLY_CONTENT[selection[0]]
        self.current_theory_item = item
        self.pending_theory_week = item["week"]
        header_line = tk.Frame(self.theory_panel, bg=COLORS["paper"])
        header_line.pack(fill="x")
        self.label(header_line, f'{item["week"]}주차 · {item["theme"]}', 16, bold=True).pack(side="left")
        self.week_list_toggle = self.action_button(header_line, "12주 목록 접기 ◀", self.toggle_theory_week_list, primary=False)
        self.week_list_toggle.pack(side="right")
        week_problems = [p for p in ALL_PROBLEMS if p["week"] == item["week"]]
        note = f'※ 필수답안 {len(item["lessons"])}개 · 전체 관련문제 {len(week_problems)}개'
        if item["week"] == 11:
            note += " · 법령·규격 수치는 최신판과 대조하세요."
        self.label(self.theory_panel, note, 9, COLORS["muted"]).pack(anchor="w", pady=(4, 10))

        notebook = ttk.Notebook(self.theory_panel)
        notebook.pack(fill="both", expand=True)
        core_tab = tk.Frame(notebook, bg=COLORS["paper"], padx=8, pady=10)
        all_tab = tk.Frame(notebook, bg=COLORS["paper"], padx=8, pady=10)
        notebook.add(core_tab, text=f"  필수답안 {len(item['lessons'])}개  ")
        notebook.add(all_tab, text=f"  전체 관련문제 {len(week_problems)}개  ")

        lesson_names = [lesson["title"] for lesson in item["lessons"]]
        core_pane = tk.PanedWindow(core_tab, orient="horizontal", bg=COLORS["paper"], sashwidth=5, sashrelief="flat")
        core_pane.pack(fill="both", expand=True)
        lesson_frame = tk.Frame(core_pane, bg=COLORS["white"], highlightbackground=COLORS["line"], highlightthickness=1)
        answer_frame = tk.Frame(core_pane, bg=COLORS["white"], highlightbackground=COLORS["line"], highlightthickness=1, padx=20, pady=16)
        core_pane.add(lesson_frame, minsize=230, width=290)
        core_pane.add(answer_frame, minsize=480)
        self.label(lesson_frame, "필수답안 목록", 10, bold=True, bg=COLORS["white"]).pack(anchor="w", padx=14, pady=(14, 8))
        self.theory_lesson_list = tk.Listbox(lesson_frame, bd=0, highlightthickness=0, selectbackground=COLORS["green"], selectforeground="white", activestyle="none", font=("맑은 고딕", 10))
        self.theory_lesson_list.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        for number, lesson_name in enumerate(lesson_names, 1):
            self.theory_lesson_list.insert("end", f"{number}. {lesson_name}")
        self.core_pane = core_pane
        self.core_lesson_frame = lesson_frame
        self.core_answer_frame = answer_frame
        core_controls = tk.Frame(answer_frame, bg=COLORS["white"])
        core_controls.pack(fill="x", pady=(0, 8))
        self.core_list_toggle = self.action_button(core_controls, "필수답안 목록 접기 ◀", self.toggle_core_list, primary=False)
        self.core_list_toggle.pack(side="right")
        self.action_button(core_controls, "답안 크게 보기", self.open_core_answer_large).pack(side="right", padx=6)
        self.theory_source = self.label(answer_frame, "", 8, COLORS["green2"], True)
        self.theory_source.pack(anchor="w")
        self.theory_title = self.label(answer_frame, "", 15, bold=True)
        self.theory_title.pack(anchor="w", pady=(3, 10))
        answer_text_frame = tk.Frame(answer_frame, bg=COLORS["white"])
        answer_text_frame.pack(fill="both", expand=True)
        self.theory_text = tk.Text(answer_text_frame, wrap="word", state="disabled", bg=COLORS["white"], fg=COLORS["ink"], relief="flat", font=("맑은 고딕", 10), padx=2, pady=2, spacing1=3, spacing3=7)
        answer_scroll = ttk.Scrollbar(answer_text_frame, orient="vertical", command=self.theory_text.yview)
        self.theory_text.configure(yscrollcommand=answer_scroll.set)
        self.theory_text.pack(side="left", fill="both", expand=True)
        answer_scroll.pack(side="right", fill="y")
        self.theory_lesson_list.bind("<<ListboxSelect>>", self.show_theory_lesson)
        self.theory_lesson_list.selection_set(0)
        self.theory_lesson_list.activate(0)
        self.show_theory_lesson()

        problem_pane = tk.PanedWindow(all_tab, orient="horizontal", bg=COLORS["paper"], sashwidth=5, sashrelief="flat")
        problem_pane.pack(fill="both", expand=True)
        tree_frame = tk.Frame(problem_pane, bg=COLORS["white"])
        detail_frame = tk.Frame(problem_pane, bg=COLORS["white"], highlightbackground=COLORS["line"], highlightthickness=1, padx=18, pady=16)
        problem_pane.add(tree_frame, minsize=355, width=430)
        problem_pane.add(detail_frame, minsize=360)
        self.theory_problem_tree = ttk.Treeview(tree_frame, show="tree", selectmode="browse")
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.theory_problem_tree.yview)
        tree_xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.theory_problem_tree.xview)
        self.theory_problem_tree.configure(yscrollcommand=tree_scroll.set, xscrollcommand=tree_xscroll.set)
        self.theory_problem_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll.grid(row=0, column=1, sticky="ns")
        tree_xscroll.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        groups = {}
        for problem in week_problems:
            group = problem["study_group"]
            if group not in groups:
                groups[group] = self.theory_problem_tree.insert("", "end", text=f"{group}", open=True)
            self.theory_problem_tree.insert(groups[group], "end", iid=f'problem-{problem["number"]}', text=f'{problem["number"]}. {problem["title"]}')
        self.theory_problem_map = {f'problem-{p["number"]}': p for p in week_problems}
        self.theory_problem_meta = self.label(detail_frame, "문제를 선택하세요.", 8, COLORS["green2"], True, wraplength=620, justify="left")
        self.theory_problem_meta.pack(anchor="w")
        self.theory_problem_title = self.label(detail_frame, "전체 관련문제", 14, bold=True, wraplength=620, justify="left")
        self.theory_problem_title.pack(anchor="w", pady=(5, 12))
        self.theory_problem_tags = self.label(detail_frame, "왼쪽에서 핵심원리와 문제를 선택하면 상세정보가 표시됩니다.", 9, COLORS["muted"], wraplength=620, justify="left")
        self.theory_problem_tags.pack(anchor="w")
        action_bar = tk.Frame(detail_frame, bg=COLORS["white"])
        action_bar.pack(anchor="w", pady=(18, 0))
        self.full_problem_pane = problem_pane
        self.full_tree_frame = tree_frame
        self.full_detail_frame = detail_frame
        self.full_list_toggle = self.action_button(action_bar, "문제 목록 접기 ◀", self.toggle_full_problem_list, primary=False)
        self.full_list_toggle.pack(side="left", padx=(0, 8))
        self.action_button(action_bar, "원문 해설 보기", self.open_selected_theory_source).pack(side="left")
        self.action_button(action_bar, "이 문제 답안연습", self.open_selected_theory_answer, primary=False).pack(side="left", padx=8)
        def resize_detail(event):
            width = max(260, event.width - 38)
            self.theory_problem_meta.config(wraplength=width)
            self.theory_problem_title.config(wraplength=width)
            self.theory_problem_tags.config(wraplength=width)
        detail_frame.bind("<Configure>", resize_detail)
        self.theory_problem_tree.bind("<<TreeviewSelect>>", self.show_theory_problem)
        first_problem = next(iter(self.theory_problem_map), None)
        if first_problem:
            self.theory_problem_tree.selection_set(first_problem)
            self.theory_problem_tree.focus(first_problem)
            self.theory_problem_tree.see(first_problem)
            self.show_theory_problem()

    def show_theory_lesson(self, *_args):
        if not hasattr(self, "theory_lesson_list") or not self.theory_lesson_list.winfo_exists():
            return
        selection = self.theory_lesson_list.curselection()
        index = selection[0] if selection else 0
        lesson = self.current_theory_item["lessons"][index]
        source_label = lesson["source"].replace("(1)", "교재")
        self.theory_source.config(text=f'교재 위치: {source_label}')
        self.theory_title.config(text=lesson["title"])
        self.theory_text.config(state="normal")
        self.theory_text.delete("1.0", "end")
        self.theory_text.insert("1.0", lesson["answer"])
        self.theory_text.see("1.0")
        self.theory_text.config(state="disabled")
        self.theory_panel.update_idletasks()

    def toggle_theory_week_list(self):
        if self.theory_selector.winfo_manager():
            self.theory_selector.pack_forget()
            self.week_list_toggle.config(text="12주 목록 펼치기 ▶")
        else:
            self.theory_selector.pack(side="left", fill="y", padx=(0, 14), before=self.theory_panel)
            self.week_list_toggle.config(text="12주 목록 접기 ◀")

    def toggle_core_list(self):
        panes = [str(pane) for pane in self.core_pane.panes()]
        lesson_path = str(self.core_lesson_frame)
        if lesson_path in panes:
            self.core_pane.forget(self.core_lesson_frame)
            self.core_list_toggle.config(text="필수답안 목록 펼치기 ▶")
        else:
            self.core_pane.add(self.core_lesson_frame, before=self.core_answer_frame, minsize=230, width=290)
            self.core_list_toggle.config(text="필수답안 목록 접기 ◀")

    def toggle_full_problem_list(self):
        panes = [str(pane) for pane in self.full_problem_pane.panes()]
        tree_path = str(self.full_tree_frame)
        if tree_path in panes:
            self.full_problem_pane.forget(self.full_tree_frame)
            self.full_list_toggle.config(text="문제 목록 펼치기 ▶")
        else:
            self.full_problem_pane.add(self.full_tree_frame, before=self.full_detail_frame, minsize=355, width=430)
            self.full_list_toggle.config(text="문제 목록 접기 ◀")

    def open_core_answer_large(self):
        title = self.theory_title.cget("text")
        source = self.theory_source.cget("text")
        content = self.theory_text.get("1.0", "end").strip()
        viewer = tk.Toplevel(self)
        viewer.title(f"필수답안 - {title}")
        viewer.geometry("1100x820")
        viewer.minsize(760, 560)
        viewer.configure(bg=COLORS["paper"])
        top = tk.Frame(viewer, bg=COLORS["green"], padx=28, pady=20)
        top.pack(fill="x")
        self.label(top, source, 9, COLORS["lime"], True, COLORS["green"]).pack(anchor="w")
        self.label(top, title, 20, "white", True, COLORS["green"], wraplength=1000, justify="left").pack(anchor="w", pady=(5, 0))
        text_frame = tk.Frame(viewer, bg=COLORS["white"])
        text_frame.pack(fill="both", expand=True, padx=24, pady=20)
        text = tk.Text(text_frame, wrap="word", bg=COLORS["white"], fg=COLORS["ink"], relief="flat", font=("맑은 고딕", 12), padx=24, pady=20, spacing1=4, spacing3=10)
        scroll = ttk.Scrollbar(text_frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.insert("1.0", content)
        text.config(state="disabled")
        text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def selected_theory_problem(self):
        selection = self.theory_problem_tree.selection() if hasattr(self, "theory_problem_tree") else ()
        return self.theory_problem_map.get(selection[0]) if selection else None

    def show_theory_problem(self, *_args):
        problem = self.selected_theory_problem()
        if not problem:
            return
        self.theory_problem_meta.config(text=f'{problem["week"]:02d}주 · {problem["study_group"]} · {problem["answer_type"]} · {problem["difficulty"]} · 교재 p.{problem["book_page"]}')
        self.theory_problem_title.config(text=f'{problem["number"]}. {problem["title"]}')
        related = " · ".join(problem["tags"])
        self.theory_problem_tags.config(text=f"연관 원리: {related}\n\n원문 해설을 열어 교재 답안을 확인하거나, 답안연습으로 이동해 직접 작성할 수 있습니다.")

    def open_selected_theory_source(self):
        problem = self.selected_theory_problem()
        if problem:
            self.open_source_page(problem)

    def open_selected_theory_answer(self):
        problem = self.selected_theory_problem()
        if problem:
            self.open_book_answer(problem)

    def render_question_list(self):
        root = self.question_scroll.content
        for child in root.winfo_children():
            child.destroy()
        term = self.search_var.get().strip().lower()
        week = self.selected_study_week()
        group = self.study_group_var.get()
        answer_kind = self.answer_type_var.get()
        visible = [q for q in ALL_PROBLEMS if (week is None or q["week"] == week) and (group == "전체 원리" or q["study_group"] == group) and (answer_kind == "전체 유형" or q["answer_type"] == answer_kind) and term in (q["title"] + " " + " ".join(q["tags"])).lower()]
        self.label(root, f"{len(visible)}개의 문제 · 유사 원리끼리 순서대로 표시", 10, COLORS["muted"]).pack(anchor="w", pady=(0, 8))
        for problem in visible:
            qid = f'book-{problem["number"]}'
            question = problem["title"]
            row = tk.Frame(root, bg=COLORS["white"], highlightbackground=COLORS["line"], highlightthickness=1, padx=14, pady=11)
            row.pack(fill="x", pady=4)
            var = tk.BooleanVar(value=qid in self.data["questions"])
            tk.Checkbutton(row, variable=var, bg=COLORS["white"], activebackground=COLORS["white"], command=lambda q=qid, v=var: self.toggle_list("questions", q, v.get(), refresh=False)).pack(side="left")
            info = tk.Frame(row, bg=COLORS["white"])
            info.pack(side="left", fill="x", expand=True, padx=8)
            page_text = f' · 교재 p.{problem["book_page"]}' if problem.get("book_page") else " · 페이지 확인 필요"
            meta = f'{problem["number"]}. {problem["week"]:02d}주 · {problem["study_group"]} · {problem["answer_type"]} · {problem["difficulty"]}{page_text}'
            self.label(info, meta, 8, COLORS["green2"], True).pack(anchor="w")
            self.label(info, question, 10, wraplength=690, justify="left").pack(anchor="w")
            if len(problem["tags"]) > 1:
                self.label(info, "연관: " + " · ".join(problem["tags"][1:]), 8, COLORS["muted"], wraplength=690, justify="left").pack(anchor="w", pady=(3, 0))
            actions = tk.Frame(row, bg=COLORS["white"])
            actions.pack(side="right")
            self.action_button(actions, "유사문제", lambda p=problem: self.show_similar_group(p), primary=False).pack(side="left", padx=(0, 5))
            self.action_button(actions, "원문 해설", lambda p=problem: self.open_source_page(p), primary=False).pack(side="left", padx=(0, 5))
            self.action_button(actions, "답안 작성", lambda p=problem: self.open_book_answer(p), primary=False).pack(side="left")

    def open_book_answer(self, problem):
        qid = f'book-{problem["number"]}'
        existing = next((q for q in QUESTIONS if q[0] == qid), None)
        if not existing:
            QUESTIONS.append((qid, problem["category"], problem["title"]))
        self.open_question(qid)

    def source_location(self, book_page: int):
        if book_page <= 145:
            return self.source_dir / "기술사 2022-1.pdf", max(0, book_page + 20)
        if book_page <= 300:
            return self.source_dir / "기술사 2022-2.pdf", max(0, book_page - 135)
        return self.source_dir / "기술사 2022-3.pdf", max(0, book_page - 289)

    def open_source_page(self, problem):
        book_page = problem.get("book_page")
        if not book_page:
            messagebox.showinfo("원문 위치", "OCR 목차에서 페이지 번호를 확정하지 못했습니다. 문제번호로 원본 목차를 확인해 주세요.")
            return
        pdf_path, page_index = self.source_location(book_page)
        if not pdf_path.exists():
            messagebox.showerror("원문 파일 없음", f"PDF를 찾을 수 없습니다.\n{pdf_path}")
            return
        viewer = tk.Toplevel(self)
        viewer.title(f'{problem["number"]}. 원문 해설 - 교재 p.{book_page}')
        viewer.geometry("980x850")
        viewer.configure(bg=COLORS["paper"])
        toolbar = tk.Frame(viewer, bg=COLORS["ink"], padx=12, pady=9)
        toolbar.pack(fill="x")
        page_label = tk.Label(toolbar, text="", bg=COLORS["ink"], fg="white", font=("맑은 고딕", 10, "bold"))
        page_label.pack(side="left", padx=10)
        canvas_frame = tk.Frame(viewer, bg="#777777")
        canvas_frame.pack(fill="both", expand=True)
        canvas = tk.Canvas(canvas_frame, bg="#777777", highlightthickness=0)
        xscroll = ttk.Scrollbar(canvas_frame, orient="horizontal", command=canvas.xview)
        yscroll = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        canvas.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)
        doc = fitz.open(pdf_path)
        state = {"index": min(page_index, doc.page_count - 1), "photo": None}
        def render():
            pix = doc[state["index"]].get_pixmap(matrix=fitz.Matrix(1.45, 1.45), alpha=False)
            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            state["photo"] = ImageTk.PhotoImage(image)
            canvas.delete("all")
            canvas.create_image(12, 12, anchor="nw", image=state["photo"])
            canvas.configure(scrollregion=(0, 0, image.width + 24, image.height + 24))
            canvas.xview_moveto(0)
            canvas.yview_moveto(0)
            page_label.config(text=f'{pdf_path.name} · PDF {state["index"] + 1}/{doc.page_count} · 교재 p.{book_page} 부근')
        def move(delta):
            state["index"] = min(max(0, state["index"] + delta), doc.page_count - 1)
            render()
        tk.Button(toolbar, text="◀ 이전", command=lambda: move(-1), bg=COLORS["lime"], fg=COLORS["ink"], bd=0, padx=12, pady=6).pack(side="right", padx=3)
        tk.Button(toolbar, text="다음 ▶", command=lambda: move(1), bg=COLORS["lime"], fg=COLORS["ink"], bd=0, padx=12, pady=6).pack(side="right", padx=3)
        viewer.protocol("WM_DELETE_WINDOW", lambda: (doc.close(), viewer.destroy()))
        render()

    def toggle_list(self, key, value, checked, refresh=False):
        items = self.data[key]
        if checked and value not in items:
            items.append(value)
        elif not checked and value in items:
            items.remove(value)
        self.save()
        if refresh:
            self.show_view("plan")

    def open_question(self, qid):
        self.pending_question = qid
        self.show_view("answer")

    def build_answer(self):
        body = tk.Frame(self.view_host, bg=COLORS["paper"])
        body.pack(fill="both", expand=True)
        guide = tk.Frame(body, bg=COLORS["soft"], padx=20, pady=20, width=220)
        guide.pack(side="left", fill="y", padx=(0, 14))
        guide.pack_propagate(False)
        self.label(guide, "답안 기본 구조", 13, bold=True, bg=COLORS["soft"]).pack(anchor="w")
        structure = "1. 정의와 개요\n2. 원리와 개념도\n3. 구성 및 절차\n4. 영향인자\n5. 장점과 한계\n6. 적용과 결론"
        self.label(guide, structure, 9, COLORS["muted"], bg=COLORS["soft"], justify="left").pack(anchor="w", pady=14)
        self.timer_label = self.label(guide, "25:00", 25, bold=True, bg=COLORS["soft"])
        self.timer_label.pack(anchor="w", pady=(20, 8))
        controls = tk.Frame(guide, bg=COLORS["soft"])
        controls.pack(anchor="w")
        self.action_button(controls, "시작/정지", self.toggle_timer, primary=False).pack(side="left")
        self.action_button(controls, "초기화", self.reset_timer, primary=False).pack(side="left", padx=5)

        editor = tk.Frame(body, bg=COLORS["white"], highlightbackground=COLORS["line"], highlightthickness=1, padx=22, pady=18)
        editor.pack(side="left", fill="both", expand=True)
        self.label(editor, "연습문제", 9, bold=True).pack(anchor="w")
        self.answer_question_var = tk.StringVar()
        display = [f"[{q[1]}] {q[2]}" for q in QUESTIONS]
        self.answer_combo = ttk.Combobox(editor, values=display, state="readonly", textvariable=self.answer_question_var)
        self.answer_combo.pack(fill="x", pady=(5, 12))
        qid = getattr(self, "pending_question", QUESTIONS[0][0])
        index = next((i for i, q in enumerate(QUESTIONS) if q[0] == qid), 0)
        self.answer_combo.current(index)
        self.answer_combo.bind("<<ComboboxSelected>>", lambda _e: self.load_selected_answer())
        self.label(editor, "답안 작성", 9, bold=True).pack(anchor="w")
        self.answer_text = tk.Text(editor, wrap="word", undo=True, font=("맑은 고딕", 10), relief="solid", bd=1, padx=12, pady=12, spacing1=3, spacing3=5)
        self.answer_text.pack(fill="both", expand=True, pady=(5, 12))
        footer = tk.Frame(editor, bg=COLORS["white"])
        footer.pack(fill="x")
        self.answer_status = self.label(footer, "", 9, COLORS["muted"])
        self.answer_status.pack(side="left")
        self.action_button(footer, "답안 저장", self.save_answer).pack(side="right")
        self.load_selected_answer()

    def selected_qid(self):
        return QUESTIONS[self.answer_combo.current()][0]

    def load_selected_answer(self):
        qid = self.selected_qid()
        self.answer_text.delete("1.0", "end")
        self.answer_text.insert("1.0", self.data["answers"].get(qid, {}).get("text", "1. 개요\n\n2. 원리\n\n3. 특징 및 적용\n"))
        self.answer_status.config(text="저장된 답안" if qid in self.data["answers"] else "새 답안")

    def save_answer(self):
        qid = self.selected_qid()
        text = self.answer_text.get("1.0", "end").rstrip()
        self.data["answers"][qid] = {"text": text, "updated": datetime.now().isoformat(timespec="seconds")}
        self.save()
        self.answer_status.config(text=f"저장 완료 · {len(text):,}자")

    def toggle_timer(self):
        self.timer_running = not self.timer_running
        if self.timer_running:
            self.tick_timer()
        elif self.timer_job:
            self.after_cancel(self.timer_job)
            self.timer_job = None

    def tick_timer(self):
        if not self.timer_running:
            return
        mins, secs = divmod(self.timer_seconds, 60)
        self.timer_label.config(text=f"{mins:02d}:{secs:02d}")
        if self.timer_seconds <= 0:
            self.timer_running = False
            messagebox.showinfo("연습 종료", "25분 답안 연습이 끝났습니다.")
            return
        self.timer_seconds -= 1
        self.timer_job = self.after(1000, self.tick_timer)

    def reset_timer(self):
        self.timer_running = False
        if self.timer_job:
            self.after_cancel(self.timer_job)
        self.timer_job = None
        self.timer_seconds = 1500
        self.timer_label.config(text="25:00")

    def build_mistakes(self):
        top = tk.Frame(self.view_host, bg=COLORS["paper"])
        top.pack(fill="x", pady=(0, 12))
        self.action_button(top, "+ 새 오답 기록", self.add_mistake_dialog).pack(side="right")
        scroll = ScrollFrame(self.view_host)
        scroll.pack(fill="both", expand=True)
        root = scroll.content
        if not self.data["mistakes"]:
            empty = self.card(root, fill="x", pady=6)
            self.label(empty, "아직 오답 기록이 없습니다.", 13, bold=True).pack(pady=(18, 3))
            self.label(empty, "공부하다 헷갈린 내용을 바로 남겨보세요.", 9, COLORS["muted"]).pack(pady=(0, 18))
            return
        for idx, item in enumerate(self.data["mistakes"]):
            row = tk.Frame(root, bg=COLORS["white"], highlightbackground=COLORS["line"], highlightthickness=1, padx=18, pady=14)
            row.pack(fill="x", pady=5)
            info = tk.Frame(row, bg=COLORS["white"])
            info.pack(side="left", fill="x", expand=True)
            self.label(info, f'{item["review_date"]} 복습', 8, COLORS["green2"], True).pack(anchor="w")
            self.label(info, item["topic"], 12, bold=True).pack(anchor="w", pady=3)
            self.label(info, item["note"], 9, COLORS["muted"], justify="left", wraplength=680).pack(anchor="w")
            tk.Button(row, text="삭제", bd=0, bg=COLORS["white"], fg=COLORS["red"], command=lambda i=idx: self.delete_mistake(i)).pack(side="right")

    def add_mistake_dialog(self):
        dialog = tk.Toplevel(self)
        dialog.title("새 오답 기록")
        dialog.geometry("520x410")
        dialog.configure(bg=COLORS["paper"])
        dialog.transient(self)
        dialog.grab_set()
        frame = tk.Frame(dialog, bg=COLORS["paper"], padx=24, pady=22)
        frame.pack(fill="both", expand=True)
        self.label(frame, "문제 또는 주제", 9, bold=True).pack(anchor="w")
        topic = tk.Entry(frame, font=("맑은 고딕", 10), relief="solid", bd=1)
        topic.pack(fill="x", ipady=7, pady=(5, 14))
        self.label(frame, "틀린 내용과 정확한 개념", 9, bold=True).pack(anchor="w")
        note = tk.Text(frame, height=8, font=("맑은 고딕", 10), relief="solid", bd=1, padx=8, pady=8)
        note.pack(fill="both", expand=True, pady=(5, 14))
        review = ttk.Combobox(frame, values=["1일 후", "7일 후", "30일 후"], state="readonly")
        review.current(1)
        review.pack(fill="x")
        def submit():
            if not topic.get().strip() or not note.get("1.0", "end").strip():
                messagebox.showwarning("입력 확인", "주제와 내용을 모두 입력하세요.", parent=dialog)
                return
            days = [1, 7, 30][review.current()]
            self.data["mistakes"].insert(0, {"topic": topic.get().strip(), "note": note.get("1.0", "end").strip(), "review_date": (date.today() + timedelta(days=days)).strftime("%Y.%m.%d")})
            self.save()
            dialog.destroy()
            self.show_view("mistakes")
        self.action_button(frame, "저장", submit).pack(anchor="e", pady=(14, 0))
        topic.focus_set()

    def delete_mistake(self, index):
        if messagebox.askyesno("삭제", "이 오답 기록을 삭제할까요?"):
            self.data["mistakes"].pop(index)
            self.save()
            self.show_view("mistakes")


if __name__ == "__main__":
    StudyApp().mainloop()
