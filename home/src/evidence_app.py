import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import datetime
from PIL import Image, ImageTk, ImageEnhance
from PIL.ExifTags import TAGS
import pandas as pd

class OverlayWindow(tk.Toplevel):
    def __init__(self, parent, bg_path, fg_path, app_ref):
        super().__init__(parent)
        self.title("지적도-사진 정밀 겹쳐보기 (오버레이 분석기)")
        self.geometry("1000x800")
        self.app_ref = app_ref
        
        self.bg_orig = None
        self.fg_orig = None
        self.display_img = None
        
        # State
        self.fg_scale = 1.0
        self.fg_x = 0
        self.fg_y = 0
        self.alpha = 0.5
        
        # Controls
        ctrl_frame = ttk.Frame(self, padding=10)
        ctrl_frame.pack(side="top", fill="x")
        
        ttk.Label(ctrl_frame, text="지적도 투명도:").pack(side="left")
        self.alpha_scale = ttk.Scale(ctrl_frame, from_=0.0, to=1.0, value=0.5, orient="horizontal", command=self.on_alpha_change)
        self.alpha_scale.pack(side="left", padx=10, fill="x", expand=True)
        
        ttk.Button(ctrl_frame, text="크기 초기화", command=self.reset_fg).pack(side="left", padx=5)
        ttk.Button(ctrl_frame, text="합성 결과 저장 (새 증거사진으로 등록)", command=self.save_composite).pack(side="left", padx=5)
        
        # Canvas
        self.cvs = tk.Canvas(self, bg="black")
        self.cvs.pack(fill="both", expand=True)
        
        self.cvs.bind("<MouseWheel>", self.zoom)
        self.cvs.bind("<ButtonPress-1>", self.scan_mark)
        self.cvs.bind("<B1-Motion>", self.scan_drag)
        
        self.load_images(bg_path, fg_path)
        
    def load_images(self, bg_path, fg_path):
        try:
            self.bg_orig = Image.open(bg_path).convert("RGBA")
            self.fg_orig = Image.open(fg_path).convert("RGBA")
            
            # Initial positioning
            self.fg_x = self.bg_orig.width // 2
            self.fg_y = self.bg_orig.height // 2
            self.fg_scale = min(self.bg_orig.width / self.fg_orig.width, self.bg_orig.height / self.fg_orig.height)
            
            # Fit canvas to window, scale bg down if too large
            self.bg_scale = 1.0
            if self.bg_orig.width > 1200 or self.bg_orig.height > 800:
                self.bg_scale = min(1200/self.bg_orig.width, 800/self.bg_orig.height)
                
            self.redraw()
        except Exception as e:
            messagebox.showerror("오류", f"이미지 로드 중 오류가 발생했습니다: {e}")
            self.destroy()
            
    def reset_fg(self):
        self.fg_x = self.bg_orig.width // 2
        self.fg_y = self.bg_orig.height // 2
        self.fg_scale = min(self.bg_orig.width / self.fg_orig.width, self.bg_orig.height / self.fg_orig.height)
        self.redraw()
        
    def on_alpha_change(self, val):
        self.alpha = float(val)
        self.redraw()
        
    def zoom(self, event):
        if event.delta > 0:
            self.fg_scale *= 1.05
        elif event.delta < 0:
            self.fg_scale /= 1.05
        self.redraw()
        
    def scan_mark(self, event):
        self.scan_mark_x = event.x
        self.scan_mark_y = event.y

    def scan_drag(self, event):
        dx = (event.x - self.scan_mark_x) / self.bg_scale
        dy = (event.y - self.scan_mark_y) / self.bg_scale
        self.fg_x += dx
        self.fg_y += dy
        self.scan_mark_x = event.x
        self.scan_mark_y = event.y
        self.redraw()
        
    def create_composite(self):
        # Create a blank image same size as bg
        comp = self.bg_orig.copy()
        
        # Resize FG
        new_w = int(self.fg_orig.width * self.fg_scale)
        new_h = int(self.fg_orig.height * self.fg_scale)
        if new_w > 0 and new_h > 0:
            fg_resized = self.fg_orig.resize((new_w, new_h), Image.Resampling.LANCZOS)
            
            # Apply Alpha
            alpha_layer = fg_resized.split()[3]
            alpha_layer = alpha_layer.point(lambda p: int(p * self.alpha))
            fg_resized.putalpha(alpha_layer)
            
            # Paste onto bg
            box_x = int(self.fg_x - new_w // 2)
            box_y = int(self.fg_y - new_h // 2)
            comp.paste(fg_resized, (box_x, box_y), fg_resized)
            
        return comp

    def redraw(self):
        if not self.bg_orig: return
        comp = self.create_composite()
        
        # scale down for display if needed
        if self.bg_scale != 1.0:
            disp_w = int(comp.width * self.bg_scale)
            disp_h = int(comp.height * self.bg_scale)
            comp = comp.resize((disp_w, disp_h), Image.Resampling.LANCZOS)
            
        self.display_img = ImageTk.PhotoImage(comp)
        self.cvs.delete("all")
        self.cvs.create_image(comp.width//2, comp.height//2, image=self.display_img, anchor="center")
        
    def save_composite(self):
        if not self.bg_orig: return
        comp = self.create_composite()
        comp = comp.convert("RGB") # Convert to save as JPG/PNG without alpha issues if needed
        
        save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"overlay_result_{int(datetime.datetime.now().timestamp())}.jpg")
        comp.save(save_path, "JPEG", quality=90)
        
        messagebox.showinfo("저장 성공", f"합성된 이미지가 저장되었습니다.\n{save_path}")
        # Automatically set as the photo_path for the current event
        if self.app_ref.current_event_id is not None:
            self.app_ref.events[self.app_ref.current_event_id]["photo_path"] = save_path
            self.app_ref.save_data()
            self.app_ref.show_photo(save_path)
        
        self.destroy()


class ZoomCanvas(tk.Canvas):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.bind("<MouseWheel>", self.zoom)
        self.bind("<ButtonPress-1>", self.scan_mark)
        self.bind("<B1-Motion>", self.scan_drag)
        self.image = None
        self.photo_img = None
        self.scale = 1.0
        self.img_x = 0
        self.img_y = 0
        self.orig_image = None
        
    def load_image(self, path):
        if not path or not os.path.exists(path):
            self.delete("all")
            self.orig_image = None
            return
        try:
            self.orig_image = Image.open(path)
            cw = self.winfo_width()
            ch = self.winfo_height()
            if cw <= 1 or ch <= 1:
                cw = 400
                ch = 400
            self.scale = min(cw/self.orig_image.width, ch/self.orig_image.height) * 0.95
            if self.scale <= 0: self.scale = 1.0
            self.img_x = cw / 2
            self.img_y = ch / 2
            self.redraw()
        except Exception as e:
            print("Image load error:", e)
            
    def redraw(self):
        if not self.orig_image:
            return
        self.delete("all")
        new_w = int(self.orig_image.width * self.scale)
        new_h = int(self.orig_image.height * self.scale)
        if new_w > 0 and new_h > 0:
            resized = self.orig_image.resize((new_w, new_h), Image.Resampling.LANCZOS)
            self.photo_img = ImageTk.PhotoImage(resized)
            self.create_image(self.img_x, self.img_y, image=self.photo_img, anchor="center")
            
    def zoom(self, event):
        if not self.orig_image: return
        if event.delta > 0:
            self.scale *= 1.2
        elif event.delta < 0:
            self.scale /= 1.2
        self.redraw()
        
    def scan_mark(self, event):
        self.scan_mark_x = event.x
        self.scan_mark_y = event.y

    def scan_drag(self, event):
        if not self.orig_image: return
        dx = event.x - self.scan_mark_x
        dy = event.y - self.scan_mark_y
        self.img_x += dx
        self.img_y += dy
        self.scan_mark_x = event.x
        self.scan_mark_y = event.y
        self.redraw()


class EvidenceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("분쟁 증거 수집기 (Evidence Tracker)")
        self.root.geometry("1100x800")
        
        self.data_file = "evidence_data.json"
        self.events = []
        self.load_data()
        
        self.create_widgets()
        
    def load_data(self):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    self.events = json.load(f)
            except:
                self.events = []
                
    def save_data(self):
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(self.events, f, ensure_ascii=False, indent=4)
            
    def create_widgets(self):
        left_frame = ttk.Frame(self.root, padding=10)
        left_frame.pack(side="left", fill="y")
        
        ttk.Label(left_frame, text="사건 타임라인", font=("Malgun Gothic", 12, "bold")).pack(pady=5)
        
        self.tree = ttk.Treeview(left_frame, columns=("Date", "Title"), show="headings", height=25)
        self.tree.heading("Date", text="일시")
        self.tree.heading("Title", text="사건 요약")
        self.tree.column("Date", width=120)
        self.tree.column("Title", width=200)
        self.tree.pack(fill="y", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select_event)
        
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="추가", command=self.add_event).pack(side="left", padx=2, expand=True, fill="x")
        ttk.Button(btn_frame, text="삭제", command=self.delete_event).pack(side="left", padx=2, expand=True, fill="x")
        
        right_frame = ttk.Frame(self.root, padding=10)
        right_frame.pack(side="right", fill="both", expand=True)
        
        land_frame = ttk.LabelFrame(right_frame, text="토지대장 기본 정보 (지적측량결과부 기준)", padding=10)
        land_frame.pack(fill="x", pady=5)
        
        ttk.Label(land_frame, text="토지 소재:").grid(row=0, column=0, sticky="w", pady=2, padx=5)
        ttk.Label(land_frame, text="서울특별시 노원구 공릉동 441-97").grid(row=0, column=1, sticky="w", pady=2, padx=5)
        
        ttk.Label(land_frame, text="지목 / 면적:").grid(row=0, column=2, sticky="w", pady=2, padx=20)
        ttk.Label(land_frame, text="대 / 147 ㎡").grid(row=0, column=3, sticky="w", pady=2, padx=5)
        
        ttk.Label(land_frame, text="측량 종목:").grid(row=1, column=0, sticky="w", pady=2, padx=5)
        ttk.Label(land_frame, text="경계복원측량 (2024.08.06)").grid(row=1, column=1, sticky="w", pady=2, padx=5)
        
        ttk.Label(land_frame, text="개별공시지가:").grid(row=1, column=2, sticky="w", pady=2, padx=20)
        ttk.Label(land_frame, text="6,468,000 원/㎡ (24.01.01 기준)").grid(row=1, column=3, sticky="w", pady=2, padx=5)
        
        ttk.Label(right_frame, text="사건 상세 기록", font=("Malgun Gothic", 12, "bold")).pack(pady=5)
        
        form_frame = ttk.Frame(right_frame)
        form_frame.pack(fill="x", pady=5)
        
        ttk.Label(form_frame, text="일시 (YYYY-MM-DD HH:MM):").grid(row=0, column=0, sticky="w", pady=2)
        self.ent_date = ttk.Entry(form_frame, width=20)
        self.ent_date.grid(row=0, column=1, sticky="w", pady=2)
        
        ttk.Label(form_frame, text="사건 요약:").grid(row=1, column=0, sticky="w", pady=2)
        self.ent_title = ttk.Entry(form_frame, width=50)
        self.ent_title.grid(row=1, column=1, sticky="w", pady=2)
        
        ttk.Label(form_frame, text="상세 내용:").grid(row=2, column=0, sticky="nw", pady=2)
        self.txt_desc = tk.Text(form_frame, width=50, height=3)
        self.txt_desc.grid(row=2, column=1, sticky="w", pady=2)
        
        ttk.Button(form_frame, text="저장", command=self.save_event_details).grid(row=3, column=1, sticky="e", pady=5)
        
        ttk.Separator(right_frame, orient="horizontal").pack(fill="x", pady=5)
        
        images_frame = ttk.Frame(right_frame)
        images_frame.pack(fill="both", expand=True, pady=5)
        images_frame.columnconfigure(0, weight=1)
        images_frame.columnconfigure(1, weight=1)
        
        photo_frame = ttk.Frame(images_frame)
        photo_frame.grid(row=0, column=0, sticky="nsew", padx=5)
        ttk.Label(photo_frame, text="증거 사진 (배경용)", font=("Malgun Gothic", 10, "bold")).pack(pady=2)
        ttk.Button(photo_frame, text="사진 추가", command=self.add_photo).pack()
        self.photo_info_lbl = ttk.Label(photo_frame, text="사진을 선택하세요 (EXIF 데이터 자동 분석)", foreground="blue")
        self.photo_info_lbl.pack(pady=2)
        self.cvs_img = ZoomCanvas(photo_frame, bg="white")
        self.cvs_img.pack(fill="both", expand=True, pady=2)
        
        map_frame = ttk.Frame(images_frame)
        map_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        ttk.Label(map_frame, text="지적도 / 참고 도면 (합성용)", font=("Malgun Gothic", 10, "bold")).pack(pady=2)
        ttk.Button(map_frame, text="도면 추가", command=self.add_map).pack()
        self.map_info_lbl = ttk.Label(map_frame, text="도면을 선택하세요", foreground="blue")
        self.map_info_lbl.pack(pady=2)
        self.cvs_map = ZoomCanvas(map_frame, bg="white")
        self.cvs_map.pack(fill="both", expand=True, pady=2)
        
        bottom_btn_frame = ttk.Frame(right_frame)
        bottom_btn_frame.pack(fill="x", pady=10)
        
        ttk.Button(bottom_btn_frame, text="🔍 지적도-사진 정밀 겹쳐보기 (오버레이)", command=self.open_overlay).pack(side="left", padx=5)
        ttk.Button(bottom_btn_frame, text="경찰/법원 제출용 리포트(Excel) 내보내기", command=self.export_excel).pack(side="right", padx=5)
        
        self.refresh_tree()
        self.current_event_id = None
        
    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, ev in enumerate(self.events):
            self.tree.insert("", "end", iid=str(i), values=(ev.get("date", ""), ev.get("title", "")))
            
    def on_select_event(self, event):
        sel = self.tree.selection()
        if not sel: return
        idx = int(sel[0])
        self.current_event_id = idx
        ev = self.events[idx]
        
        self.ent_date.delete(0, tk.END)
        self.ent_date.insert(0, ev.get("date", ""))
        self.ent_title.delete(0, tk.END)
        self.ent_title.insert(0, ev.get("title", ""))
        self.txt_desc.delete("1.0", tk.END)
        self.txt_desc.insert("1.0", ev.get("desc", ""))
        
        self.root.after(50, lambda: self.show_photo(ev.get("photo_path", "")))
        self.root.after(50, lambda: self.show_map(ev.get("map_path", "")))
        
    def add_event(self):
        new_ev = {
            "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "title": "새 사건",
            "desc": "",
            "photo_path": "",
            "map_path": ""
        }
        self.events.append(new_ev)
        self.save_data()
        self.refresh_tree()
        self.tree.selection_set(str(len(self.events)-1))
        
    def delete_event(self):
        if self.current_event_id is not None:
            if messagebox.askyesno("삭제", "선택한 사건을 삭제하시겠습니까?"):
                self.events.pop(self.current_event_id)
                self.save_data()
                self.refresh_tree()
                self.current_event_id = None
                self.ent_date.delete(0, tk.END)
                self.ent_title.delete(0, tk.END)
                self.txt_desc.delete("1.0", tk.END)
                self.cvs_img.delete("all")
                self.photo_info_lbl.config(text="")
                self.cvs_map.delete("all")
                self.map_info_lbl.config(text="")
        
    def save_event_details(self):
        if self.current_event_id is not None:
            ev = self.events[self.current_event_id]
            ev["date"] = self.ent_date.get()
            ev["title"] = self.ent_title.get()
            ev["desc"] = self.txt_desc.get("1.0", tk.END).strip()
            self.save_data()
            self.refresh_tree()
            messagebox.showinfo("저장", "저장되었습니다.")
            
    def add_photo(self):
        if self.current_event_id is None:
            messagebox.showwarning("경고", "먼저 타임라인에서 사건을 선택하세요.")
            return
            
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
        if file_path:
            self.events[self.current_event_id]["photo_path"] = file_path
            self.save_data()
            self.show_photo(file_path)
            
    def add_map(self):
        if self.current_event_id is None:
            messagebox.showwarning("경고", "먼저 타임라인에서 사건을 선택하세요.")
            return
            
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
        if file_path:
            self.events[self.current_event_id]["map_path"] = file_path
            self.save_data()
            self.show_map(file_path)
            
    def show_photo(self, path):
        if not path or not os.path.exists(path):
            self.cvs_img.delete("all")
            self.photo_info_lbl.config(text="등록된 사진이 없습니다.")
            return
        try:
            img = Image.open(path)
            exif_data = img._getexif()
            date_taken = "메타데이터(촬영일자) 없음"
            if exif_data:
                for tag_id, value in exif_data.items():
                    tag = TAGS.get(tag_id, tag_id)
                    if tag == 'DateTimeOriginal':
                        date_taken = f"촬영 일자: {value}"
                        break
            self.photo_info_lbl.config(text=f"증거 분석: {date_taken} | 파일: {os.path.basename(path)}")
            self.cvs_img.load_image(path)
        except Exception as e:
            self.cvs_img.delete("all")
            self.photo_info_lbl.config(text=str(e))
            
    def show_map(self, path):
        if not path or not os.path.exists(path):
            self.cvs_map.delete("all")
            self.map_info_lbl.config(text="등록된 도면이 없습니다.")
            return
        try:
            self.map_info_lbl.config(text=f"도면: {os.path.basename(path)}")
            self.cvs_map.load_image(path)
        except Exception as e:
            self.cvs_map.delete("all")
            self.map_info_lbl.config(text=str(e))
            
    def open_overlay(self):
        if self.current_event_id is None:
            messagebox.showwarning("경고", "사건을 먼저 선택하세요.")
            return
        ev = self.events[self.current_event_id]
        if not ev.get("photo_path") or not ev.get("map_path") or not os.path.exists(ev.get("photo_path")) or not os.path.exists(ev.get("map_path")):
            messagebox.showwarning("경고", "사진과 도면이 모두 등록되어 있어야 합니다.")
            return
            
        OverlayWindow(self.root, ev.get("photo_path"), ev.get("map_path"), self)
            
    def export_excel(self):
        if not self.events:
            messagebox.showinfo("알림", "출력할 데이터가 없습니다.")
            return
            
        df = pd.DataFrame(self.events)
        out_file = "증거제출용_리포트.xlsx"
        df.to_excel(out_file, index=False, columns=["date", "title", "desc", "photo_path", "map_path"])
        messagebox.showinfo("완료", f"{out_file} 파일이 생성되었습니다.\n(경찰서 및 변호사 제출용으로 활용하세요)")

if __name__ == "__main__":
    root = tk.Tk()
    app = EvidenceApp(root)
    root.mainloop()
