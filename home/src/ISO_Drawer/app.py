from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from core import Point, Project, export_dxf, export_pdf, load_project, save_project, snap_iso


class IsoDrawer(tk.Tk):
    BG, GRID, PIPE, PREVIEW = "#17232d", "#263743", "#64e6db", "#ffbf69"

    def __init__(self):
        super().__init__()
        self.title("ISO Drawer - 독립형 배관 아이소메트릭")
        self.geometry("1280x780")
        self.minsize(900, 600)
        self.project = Project()
        self.cursor = None
        self.selected_seg = None
        self.chain_breaks: set[int] = set()
        # 드래그 상태
        self._drag_pt: int | None = None
        self._drag_moved = False
        self._press_pos = (0, 0)
        self.edit_len_var = tk.StringVar()
        self._build()
        self._setup_binds()
        self.redraw()

    def _build(self):
        bar = ttk.Frame(self, padding=6); bar.pack(fill="x")
        for text, cmd in [("새 도면", self.new), ("열기", self.open), ("저장", self.save), ("DXF 출력", self.dxf), ("PDF 출력", self.pdf), ("실행 취소", self.undo), ("전체 맞춤", self.fit)]:
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=2)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=6, pady=2)
        self.snap_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="각도 스냅", variable=self.snap_var, command=self.redraw).pack(side="left", padx=2)
        ttk.Label(bar, text="  좌클릭: 점 입력 · 우클릭/ESC: 선 종료 · 스냅 OFF=자유각도 · 휠: 확대/축소 · 중간 드래그: 이동").pack(side="left")
        body = ttk.Panedwindow(self, orient="horizontal"); body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(body, bg=self.BG, highlightthickness=0, cursor="crosshair")
        panel = ttk.Frame(body, padding=12, width=230); body.add(self.canvas, weight=5); body.add(panel, weight=1)
        self.vars = {k: tk.StringVar(value=v) for k, v in {"company":"SITCO", "project_name":"", "line_no":"LINE-001", "size":'4"', "spec":"", "component":"NONE"}.items()}
        for label, key in [("회사명", "company"), ("공사명", "project_name"), ("라인 번호", "line_no"), ("배관 구경", "size"), ("SPEC", "spec")]:
            ttk.Label(panel, text=label).pack(anchor="w", pady=(8,2)); ttk.Entry(panel, textvariable=self.vars[key]).pack(fill="x")
        ttk.Label(panel, text="현재 끝점 부속").pack(anchor="w", pady=(16,2))
        ttk.Combobox(panel, textvariable=self.vars["component"], state="readonly", values=("NONE","ELBOW_90","ELBOW_45","TEE","GATE_VALVE","BALL_VALVE","CHECK_VALVE","FLANGE","REDUCER","WELD")).pack(fill="x")
        ttk.Button(panel, text="끝점에 부속 적용", command=self.apply_component).pack(fill="x", pady=6)
        ttk.Separator(panel).pack(fill="x", pady=8)
        ttk.Label(panel, text="선분 수정", font=("맑은 고딕", 10, "bold")).pack(anchor="w")
        ttk.Label(panel, text="길이(mm)").pack(anchor="w", pady=(6,2))
        ttk.Entry(panel, textvariable=self.edit_len_var).pack(fill="x")
        ttk.Button(panel, text="길이 적용", command=self.apply_seg_length).pack(fill="x", pady=3)
        ttk.Button(panel, text="선분 삭제  [Del]", command=self.delete_seg).pack(fill="x")
        ttk.Separator(panel).pack(fill="x", pady=8)
        ttk.Label(panel, text="작업 방법", font=("맑은 고딕", 10, "bold")).pack(anchor="w")
        ttk.Label(panel, justify="left", wraplength=205, text="• 좌클릭: 점 입력/연결\n• 우클릭/ESC: 선 종료\n• 점 드래그: 위치 이동\n• 선분 클릭: 선택\n• Del: 선택선분 삭제").pack(anchor="w", pady=4)
        self.status = tk.StringVar(); ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w", padding=4).pack(fill="x")
        self.scale, self.ox, self.oy = 1.0, 0.0, 0.0

    def _setup_binds(self):
        self.canvas.bind("<ButtonPress-1>", self._press)
        self.canvas.bind("<B1-Motion>", self._b1_motion)
        self.canvas.bind("<ButtonRelease-1>", self._release)
        self.canvas.bind("<Motion>", self.motion)
        self.canvas.bind("<Button-3>", lambda e: self.end_line())
        self.bind("<Escape>", lambda e: self.end_line())
        self.canvas.bind("<MouseWheel>", self.zoom)
        self.canvas.bind("<ButtonPress-2>", self.pan_start)
        self.canvas.bind("<B2-Motion>", self.pan)
        self.bind("<Control-z>", lambda e: self.undo())
        self.bind("<Delete>", lambda e: self.delete_seg())

    def world(self, x, y): return (x-self.ox)/self.scale, (y-self.oy)/self.scale
    def screen(self, x, y): return x*self.scale+self.ox, y*self.scale+self.oy

    # ── 마우스 프레스/드래그/릴리즈 ──────────────────────────────
    def _press(self, e):
        wx, wy = self.world(e.x, e.y)
        self._drag_pt = self._find_point(wx, wy)
        self._drag_moved = False
        self._press_pos = (e.x, e.y)

    def _b1_motion(self, e):
        if self._drag_pt is None: return
        dx, dy = e.x - self._press_pos[0], e.y - self._press_pos[1]
        if not self._drag_moved and (dx**2 + dy**2) < 16: return  # 4px 임계값
        self._drag_moved = True
        wx, wy = self.world(e.x, e.y)
        pts = self.project.points
        pts[self._drag_pt].x = wx
        pts[self._drag_pt].y = wy
        self.cursor = None
        self.redraw()

    def _release(self, e):
        if not self._drag_moved:
            self.click(e)  # 드래그 없으면 일반 클릭으로 전달
        self._drag_pt = None
        self._drag_moved = False

    def _find_point(self, wx, wy, tol_screen=12):
        """클릭 좌표 근처의 기존 포인트 인덱스 반환"""
        pts = self.project.points
        tol = tol_screen / self.scale
        for i, p in enumerate(pts):
            if ((wx-p.x)**2 + (wy-p.y)**2)**0.5 < tol:
                return i
        return None

    def _find_segment(self, wx, wy):
        """클릭 좌표 근처의 선분 인덱스 반환 (없으면 None, 체인 끊김 구간 제외)"""
        pts = self.project.points
        tol = 8 / self.scale
        for i, (a, b) in enumerate(zip(pts, pts[1:])):
            if (i + 1) in self.chain_breaks: continue  # 체인 끊김 구간 제외
            dx, dy = b.x - a.x, b.y - a.y
            seg_len2 = dx*dx + dy*dy
            if seg_len2 == 0: continue
            t = max(0.0, min(1.0, ((wx-a.x)*dx + (wy-a.y)*dy) / seg_len2))
            px, py = a.x + t*dx, a.y + t*dy
            if ((wx-px)**2 + (wy-py)**2)**0.5 < tol:
                return i
        return None

    def click(self, e):
        wx, wy = self.world(e.x, e.y)
        n = len(self.project.points)
        chain_active = n > 0 and n not in self.chain_breaks
        pts = self.project.points
        pos = (wx, wy)

        if not chain_active:
            # ─── 대기 모드 ───────────────────────────────────────────
            # ① 기존 점 근처 클릭 → 그 점에서 이어서 그리기 (최우선)
            pi = self._find_point(wx, wy)
            if pi is not None:
                pts.append(Point(pts[pi].x, pts[pi].y))
                self.redraw(); return
            # ② 선분 클릭 → 선택/해제
            seg = self._find_segment(wx, wy)
            if seg is not None:
                self.selected_seg = None if self.selected_seg == seg else seg
                self.redraw(); return
            # ③ 빈 공간 클릭 → 새 시작점
            self.selected_seg = None
            pts.append(Point(*pos)); self.redraw(); return

        # ─── 그리기 활성 모드 ────────────────────────────────────────
        self.selected_seg = None
        start = pts[-1]
        if self.snap_var.get():
            x, y = snap_iso((start.x, start.y), pos)
        else:
            x, y = pos
        if ((x-start.x)**2+(y-start.y)**2)**.5 < 5/self.scale: return
        raw = simpledialog.askstring("실제 길이", "이 구간의 실제 길이(mm):\n(생략 시 비워두고 OK)", parent=self)
        if raw is None: return
        try:
            length = float(raw.strip()) if raw.strip() else 0.0
        except ValueError:
            length = 0.0
        pts.append(Point(x, y, length)); self.redraw()

    def motion(self, e):
        self.cursor = self.world(e.x,e.y); self.redraw()

    def redraw(self):
        c=self.canvas; c.delete("all"); w=max(c.winfo_width(),1); h=max(c.winfo_height(),1)
        step=50*self.scale
        if step>=12:
            x=self.ox%step
            while x<w: c.create_line(x,0,x,h,fill=self.GRID); x+=step
            y=self.oy%step
            while y<h: c.create_line(0,y,w,y,fill=self.GRID); y+=step
        pts=self.project.points
        n=len(pts)
        for i,(a,b) in enumerate(zip(pts,pts[1:]),1):
            if i in self.chain_breaks: continue
            ax,ay=self.screen(a.x,a.y); bx,by=self.screen(b.x,b.y)
            selected = (self.selected_seg == i-1)
            c.create_line(ax,ay,bx,by,fill="#ffffff" if selected else self.PIPE,width=5 if selected else 3)
            if b.actual_length:
                # 선에 수직 방향으로 치수 텍스트 배치
                ldx, ldy = bx-ax, by-ay
                llen = max((ldx**2+ldy**2)**0.5, 1)
                nx, ny = -ldy/llen*10, ldx/llen*10
                if ny > 0: nx, ny = -nx, -ny  # 항상 선 위쪽에 배치
                tx, ty = (ax+bx)/2+nx, (ay+by)/2+ny
                c.create_text(tx,ty,text=f"{b.actual_length:g} mm",fill="#ffd166" if selected else "white",font=("Segoe UI",9,"bold" if selected else "normal"))
        for i,p in enumerate(pts):
            color = "#ff7b7b" if i not in self.chain_breaks else "#aaaaaa"
            x,y=self.screen(p.x,p.y); c.create_oval(x-4,y-4,x+4,y+4,outline=color,width=2)
            if p.component!="NONE": c.create_text(x+7,y+8,text=p.component,fill="#ffd166",anchor="nw",font=("Segoe UI",9,"bold"))
        chain_active = n > 0 and n not in self.chain_breaks
        if chain_active and self.cursor:
            p=pts[-1]
            q=snap_iso((p.x,p.y),self.cursor) if self.snap_var.get() else self.cursor
            a=self.screen(p.x,p.y); b=self.screen(*q)
            c.create_line(*a,*b,fill=self.PREVIEW,width=2,dash=(6,4))
        segs = max(0, n-1) - len(self.chain_breaks)
        snap_txt = "스냅 ON" if self.snap_var.get() else "자유 각도"
        chain_txt = "그리기 중" if chain_active else "대기 (클릭으로 새 선 시작)"
        sel_txt = f" · 선분{self.selected_seg+1} 선택중" if self.selected_seg is not None else ""
        self.status.set(f"포인트 {n}개 · 구간 {segs}개 · {snap_txt} · {chain_txt}{sel_txt}")
        # 선택된 선분 길이 자동 반영
        if self.selected_seg is not None and self.selected_seg + 1 < len(pts):
            self.edit_len_var.set(str(pts[self.selected_seg + 1].actual_length))
        elif self.selected_seg is None:
            self.edit_len_var.set("")

    def sync(self):
        self.project.company=self.vars["company"].get(); self.project.project_name=self.vars["project_name"].get()
        self.project.line_no=self.vars["line_no"].get(); self.project.size=self.vars["size"].get(); self.project.spec=self.vars["spec"].get()
    def new(self):
        if messagebox.askyesno("새 도면", "현재 도면을 지우고 새로 시작할까요?"):
            self.project=Project(); self.cursor=None; self.chain_breaks.clear(); self.selected_seg=None; self.redraw()
    def undo(self):
        if not self.project.points: return
        n = len(self.project.points)
        self.chain_breaks.discard(n)
        self.chain_breaks.discard(n-1)
        self.project.points.pop()
        self.selected_seg = None
        self.redraw()
    def apply_seg_length(self):
        """선택된 선분의 길이 수정"""
        if self.selected_seg is None:
            return messagebox.showinfo("선분 수정", "먼저 선분을 클릭해서 선택하세요.")
        try:
            length = float(self.edit_len_var.get().strip() or "0")
        except ValueError:
            return messagebox.showwarning("입력 오류", "숫자를 입력하세요.")
        self.project.points[self.selected_seg + 1].actual_length = length
        self.redraw()
    def delete_seg(self):
        """선택된 선분의 끝점 제거 (선분 삭제)"""
        if self.selected_seg is None: return
        idx = self.selected_seg + 1  # 제거할 끝점 인덱스
        pts = self.project.points
        # chain_breaks 인덱스 조정
        self.chain_breaks = {b - 1 if b > idx else b
                             for b in self.chain_breaks if b != idx}
        pts.pop(idx)
        self.selected_seg = None
        self.redraw()
    def end_line(self):
        n = len(self.project.points)
        if n > 0: self.chain_breaks.add(n)  # 다음 클릭은 새 체인 시작
        self.cursor=None; self.redraw()
    def apply_component(self):
        if not self.project.points: return
        self.project.points[-1].component=self.vars["component"].get(); self.redraw()
    def save(self):
        self.sync(); path=filedialog.asksaveasfilename(defaultextension=".json",filetypes=[("ISO 프로젝트","*.json")])
        if path: save_project(self.project,path)
    def open(self):
        path=filedialog.askopenfilename(filetypes=[("ISO 프로젝트","*.json")])
        if path:
            try:
                self.project=load_project(path)
                for k in ("company","project_name","line_no","size","spec"): self.vars[k].set(getattr(self.project,k))
                self.fit()
            except Exception as ex: messagebox.showerror("열기 실패",str(ex))
    def dxf(self): self._do_export(".dxf", export_dxf, [("DXF 도면","*.dxf")])
    def pdf(self): self._do_export(".pdf", export_pdf, [("PDF 도면","*.pdf")])
    def _do_export(self, ext, func, types):
        if len(self.project.points)<2: return messagebox.showwarning("출력 불가","두 개 이상의 포인트를 입력하세요.")
        self.sync(); path=filedialog.asksaveasfilename(defaultextension=ext,filetypes=types)
        if path:
            try: func(self.project,path); messagebox.showinfo("출력 완료",path)
            except Exception as ex: messagebox.showerror("출력 실패",str(ex))
    def zoom(self,e):
        factor=1.15 if e.delta>0 else 1/1.15; wx,wy=self.world(e.x,e.y); self.scale*=factor; self.ox=e.x-wx*self.scale; self.oy=e.y-wy*self.scale; self.redraw()
    def pan_start(self,e): self._pan=(e.x,e.y,self.ox,self.oy)
    def pan(self,e): _,_,ox,oy=self._pan; self.ox=ox+e.x-self._pan[0]; self.oy=oy+e.y-self._pan[1]; self.redraw()
    def fit(self):
        if not self.project.points: self.scale=1; self.ox=self.oy=0
        else:
            xs=[p.x for p in self.project.points]; ys=[p.y for p in self.project.points]; w=max(self.canvas.winfo_width(),600); h=max(self.canvas.winfo_height(),400)
            self.scale=min((w-120)/max(max(xs)-min(xs),1),(h-120)/max(max(ys)-min(ys),1)); self.ox=60-min(xs)*self.scale; self.oy=60-min(ys)*self.scale
        self.redraw()


if __name__ == "__main__": IsoDrawer().mainloop()
