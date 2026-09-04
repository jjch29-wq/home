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
        self._build()
        self._bind()
        self.redraw()

    def _build(self):
        bar = ttk.Frame(self, padding=6); bar.pack(fill="x")
        for text, cmd in [("새 도면", self.new), ("열기", self.open), ("저장", self.save), ("DXF 출력", self.dxf), ("PDF 출력", self.pdf), ("실행 취소", self.undo), ("전체 맞춤", self.fit)]:
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=2)
        ttk.Label(bar, text="  좌클릭: 점 입력 · 우클릭/ESC: 선 종료 · 휠: 확대/축소 · 중간 드래그: 이동").pack(side="left")
        body = ttk.Panedwindow(self, orient="horizontal"); body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(body, bg=self.BG, highlightthickness=0, cursor="crosshair")
        panel = ttk.Frame(body, padding=12, width=230); body.add(self.canvas, weight=5); body.add(panel, weight=1)
        self.vars = {k: tk.StringVar(value=v) for k, v in {"line_no":"LINE-001", "size":'4"', "spec":"", "component":"NONE"}.items()}
        for label, key in [("라인 번호", "line_no"), ("배관 구경", "size"), ("SPEC", "spec")]:
            ttk.Label(panel, text=label).pack(anchor="w", pady=(8,2)); ttk.Entry(panel, textvariable=self.vars[key]).pack(fill="x")
        ttk.Label(panel, text="현재 끝점 부속").pack(anchor="w", pady=(16,2))
        ttk.Combobox(panel, textvariable=self.vars["component"], state="readonly", values=("NONE","ELBOW_90","ELBOW_45","TEE","GATE_VALVE","BALL_VALVE","CHECK_VALVE","FLANGE","REDUCER","WELD")).pack(fill="x")
        ttk.Button(panel, text="끝점에 부속 적용", command=self.apply_component).pack(fill="x", pady=6)
        ttk.Separator(panel).pack(fill="x", pady=12)
        ttk.Label(panel, text="작업 방법", font=("맑은 고딕", 10, "bold")).pack(anchor="w")
        ttk.Label(panel, justify="left", wraplength=205, text="1. 시작점을 클릭합니다.\n2. 다음 방향을 클릭하면 ISO 각도로 고정됩니다.\n3. 실제 배관 길이(mm)를 입력합니다.\n4. 필요할 때 끝점 부속을 적용합니다.\n5. 우클릭으로 한 라인을 종료합니다.").pack(anchor="w", pady=6)
        self.status = tk.StringVar(); ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w", padding=4).pack(fill="x")
        self.scale, self.ox, self.oy = 1.0, 0.0, 0.0

    def _bind(self):
        self.canvas.bind("<Button-1>", self.click); self.canvas.bind("<Motion>", self.motion)
        self.canvas.bind("<Button-3>", lambda e: self.end_line()); self.bind("<Escape>", lambda e: self.end_line())
        self.canvas.bind("<MouseWheel>", self.zoom); self.canvas.bind("<ButtonPress-2>", self.pan_start); self.canvas.bind("<B2-Motion>", self.pan)
        self.bind("<Control-z>", lambda e: self.undo())

    def world(self, x, y): return (x-self.ox)/self.scale, (y-self.oy)/self.scale
    def screen(self, x, y): return x*self.scale+self.ox, y*self.scale+self.oy

    def click(self, e):
        pos = self.world(e.x, e.y)
        if not self.project.points:
            self.project.points.append(Point(*pos)); self.redraw(); return
        start = self.project.points[-1]; x, y = snap_iso((start.x,start.y), pos)
        if ((x-start.x)**2+(y-start.y)**2)**.5 < 5/self.scale: return
        length = simpledialog.askfloat("실제 길이", "이 구간의 실제 길이(mm):", parent=self, minvalue=0.1)
        if length is None: return
        self.project.points.append(Point(x, y, length)); self.redraw()

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
        for i,(a,b) in enumerate(zip(pts,pts[1:]),1):
            ax,ay=self.screen(a.x,a.y); bx,by=self.screen(b.x,b.y); c.create_line(ax,ay,bx,by,fill=self.PIPE,width=3)
            c.create_text((ax+bx)/2,(ay+by)/2-12,text=f"{b.actual_length:g} mm",fill="white",font=("Segoe UI",9))
        for i,p in enumerate(pts):
            x,y=self.screen(p.x,p.y); c.create_oval(x-4,y-4,x+4,y+4,outline="#ff7b7b",width=2)
            if p.component!="NONE": c.create_text(x+7,y+8,text=p.component,fill="#ffd166",anchor="nw",font=("Segoe UI",9,"bold"))
        if pts and self.cursor:
            p=pts[-1]; q=snap_iso((p.x,p.y),self.cursor); a=self.screen(p.x,p.y); b=self.screen(*q); c.create_line(*a,*b,fill=self.PREVIEW,width=2,dash=(6,4))
        self.status.set(f"포인트 {len(pts)}개 · 구간 {max(0,len(pts)-1)}개")

    def sync(self):
        self.project.line_no=self.vars["line_no"].get(); self.project.size=self.vars["size"].get(); self.project.spec=self.vars["spec"].get()
    def new(self):
        if messagebox.askyesno("새 도면", "현재 도면을 지우고 새로 시작할까요?"): self.project=Project(); self.cursor=None; self.redraw()
    def undo(self):
        if self.project.points: self.project.points.pop(); self.redraw()
    def end_line(self): self.cursor=None; self.redraw()
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
                for k in ("line_no","size","spec"): self.vars[k].set(getattr(self.project,k))
                self.fit()
            except Exception as ex: messagebox.showerror("열기 실패",str(ex))
    def dxf(self): self._export(".dxf", export_dxf, [("DXF 도면","*.dxf")])
    def pdf(self): self._export(".pdf", export_pdf, [("PDF 도면","*.pdf")])
    def _export(self, ext, func, types):
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
