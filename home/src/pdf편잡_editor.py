import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import os

class PDFEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 페이지 편집기 (PRO)")
        self.root.geometry("1200x900")
        try:
            self.root.state('zoomed') # 윈도우 최대화
        except:
            pass
        
        # State variables
        self.pages = [] # List of dicts: {'doc': fitz.Document, 'page_num': int, 'label': str, 'rotation': int}
        self.docs = []  # Keep references to memory-loaded documents
        
        self.drag_start_idx = -1
        
        self.setup_ui()
        
    def setup_ui(self):
        # Left frame: Listbox and controls
        left_frame = ttk.Frame(self.root, padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="PDF 열기", command=self.open_pdf).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="PDF 추가", command=self.add_pdf).pack(side=tk.LEFT, padx=2)
        
        # Batch Selection Buttons
        sel_frame = ttk.Frame(left_frame)
        sel_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(sel_frame, text="전체", command=lambda: self.batch_select('all'), width=5).pack(side=tk.LEFT, padx=1)
        ttk.Button(sel_frame, text="홀수", command=lambda: self.batch_select('odd'), width=5).pack(side=tk.LEFT, padx=1)
        ttk.Button(sel_frame, text="짝수", command=lambda: self.batch_select('even'), width=5).pack(side=tk.LEFT, padx=1)
        
        ttk.Label(left_frame, text="* 드래그 앤 드롭으로 순서 변경 가능", font=("Arial", 8)).pack(anchor=tk.W, pady=(5,0))
        
        # Listbox for pages
        self.listbox = tk.Listbox(left_frame, selectmode=tk.EXTENDED, width=45)
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=2)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)
        
        # Drag and drop bindings
        self.listbox.bind('<Button-1>', self.on_drag_start)
        self.listbox.bind('<B1-Motion>', self.on_drag_motion)
        
        ctrl_frame = ttk.Frame(left_frame)
        ctrl_frame.pack(fill=tk.X, pady=5)
        ttk.Button(ctrl_frame, text="▲", width=3, command=lambda: self.move_selected(-1)).pack(side=tk.LEFT, padx=1)
        ttk.Button(ctrl_frame, text="▼", width=3, command=lambda: self.move_selected(1)).pack(side=tk.LEFT, padx=1)
        ttk.Button(ctrl_frame, text="선택 삭제", command=self.delete_selected).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        
        ttk.Button(left_frame, text="저장하기", command=self.save_pdf).pack(fill=tk.X, pady=5)
        
        # Right frame: Preview and Rotation
        right_frame = ttk.Frame(self.root, padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        rot_frame = ttk.Frame(right_frame)
        rot_frame.pack(fill=tk.X, pady=5)
        ttk.Button(rot_frame, text="↺ 왼쪽으로 회전", command=lambda: self.rotate_selected(-90)).pack(side=tk.LEFT, padx=5)
        ttk.Button(rot_frame, text="↻ 오른쪽으로 회전", command=lambda: self.rotate_selected(90)).pack(side=tk.LEFT, padx=5)
        
        # Scrollable Canvas for preview
        self.preview_canvas = tk.Canvas(right_frame, bg='lightgray')
        v_scroll = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        h_scroll = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=self.preview_canvas.xview)
        self.preview_canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        # 확대/축소 마우스 휠 바인딩
        self.preview_canvas.bind('<MouseWheel>', self.on_mouse_wheel)
        self.preview_canvas.bind('<Control-MouseWheel>', self.on_mouse_wheel)
        self.preview_canvas.bind('<Button-4>', self.on_mouse_wheel) # Linux
        self.preview_canvas.bind('<Button-5>', self.on_mouse_wheel) # Linux
        self.preview_canvas.bind('<Control-Button-4>', self.on_mouse_wheel)
        self.preview_canvas.bind('<Control-Button-5>', self.on_mouse_wheel)
        
        # 마우스가 캔버스 위로 올라갈 때 포커스를 주어 휠 이벤트가 작동하도록 함
        self.preview_canvas.bind('<Enter>', lambda e: self.preview_canvas.focus_set())
        
        self.zoom_factor = 1.0
        
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 캔버스 크기 변경 시 중앙 정렬 유지를 위한 이벤트 바인딩
        self.preview_canvas.bind('<Configure>', self.on_canvas_configure)
        
        self.preview_image_id = self.preview_canvas.create_image(0, 0, anchor=tk.CENTER)
        self.preview_text_id = self.preview_canvas.create_text(300, 300, text="페이지를 선택하면 미리보기가 표시됩니다.", font=("Arial", 12), anchor=tk.CENTER)
        
        self.current_image = None
        
    def _update_preview_canvas(self, img, text):
        canvas_w = self.preview_canvas.winfo_width()
        canvas_h = self.preview_canvas.winfo_height()
        
        # 처음 렌더링 전에는 1을 반환하므로 기본값 설정
        if canvas_w < 10: canvas_w = 800
        if canvas_h < 10: canvas_h = 600
        
        if img:
            img_w = img.width()
            img_h = img.height()
            
            # 스크롤 영역은 캔버스 크기와 이미지 크기 중 큰 값
            scroll_w = max(canvas_w, img_w)
            scroll_h = max(canvas_h, img_h)
            
            x, y = scroll_w / 2, scroll_h / 2
            
            self.preview_canvas.coords(self.preview_image_id, x, y)
            self.preview_canvas.itemconfig(self.preview_image_id, image=img)
            self.preview_canvas.itemconfig(self.preview_text_id, text="")
            self.preview_canvas.config(scrollregion=(0, 0, scroll_w, scroll_h))
        else:
            x, y = canvas_w / 2, canvas_h / 2
            
            self.preview_canvas.itemconfig(self.preview_image_id, image="")
            self.preview_canvas.coords(self.preview_text_id, x, y)
            self.preview_canvas.itemconfig(self.preview_text_id, text=text)
            self.preview_canvas.config(scrollregion=(0, 0, canvas_w, canvas_h))

    def on_canvas_configure(self, event):
        # 윈도우 크기가 변경될 때 중앙 정렬을 다시 계산
        self._update_preview_canvas(self.current_image, "페이지를 선택하면 미리보기가 표시됩니다." if not getattr(self, 'current_image', None) else "")
            
    def batch_select(self, mode):
        self.listbox.selection_clear(0, tk.END)
        for i in range(len(self.pages)):
            if mode == 'all':
                self.listbox.selection_set(i)
            elif mode == 'odd' and (i % 2 == 0): # 1st page is index 0
                self.listbox.selection_set(i)
            elif mode == 'even' and (i % 2 == 1): # 2nd page is index 1
                self.listbox.selection_set(i)
        self.on_select(None)
        
    def on_drag_start(self, event):
        # Only start drag if we click on a valid item
        if self.listbox.size() > 0:
            self.drag_start_idx = self.listbox.nearest(event.y)
        else:
            self.drag_start_idx = -1
            
    def on_drag_motion(self, event):
        if self.drag_start_idx < 0: return
        idx = self.listbox.nearest(event.y)
        if idx != self.drag_start_idx and 0 <= idx < len(self.pages):
            # Swap in UI
            text = self.listbox.get(self.drag_start_idx)
            self.listbox.delete(self.drag_start_idx)
            self.listbox.insert(idx, text)
            # Swap in memory
            self.pages[self.drag_start_idx], self.pages[idx] = self.pages[idx], self.pages[self.drag_start_idx]
            
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(idx)
            self.drag_start_idx = idx
            
    def open_pdf(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if not filepaths: return
        self.pages.clear()
        self.docs.clear()
        self.listbox.delete(0, tk.END)
        self._update_preview_canvas(None, "페이지를 선택하면 미리보기가 표시됩니다.")
        for filepath in filepaths:
            self.add_document(filepath)
        
    def add_pdf(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if not filepaths: return
        
        # 선택된 항목이 있으면 그 바로 아래에, 없으면 맨 끝에 추가
        selection = self.listbox.curselection()
        insert_idx = selection[-1] + 1 if selection else len(self.pages)
        
        for filepath in filepaths:
            insert_idx = self.add_document(filepath, insert_idx)
        
    def add_document(self, filepath, insert_idx=None):
        if insert_idx is None:
            insert_idx = len(self.pages)
            
        try:
            # Memory load to release file lock immediately
            with open(filepath, "rb") as f:
                pdf_bytes = f.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            self.docs.append(doc)
            
            filename = os.path.basename(filepath)
            for i in range(len(doc)):
                label = f"{filename} - {i+1}쪽"
                self.pages.insert(insert_idx, {'doc': doc, 'page_num': i, 'label': label, 'rotation': 0})
                self.listbox.insert(insert_idx, label)
                insert_idx += 1
                
            return insert_idx
        except Exception as e:
            messagebox.showerror("오류", f"PDF를 열 수 없습니다:\n{e}")
            
    def update_preview(self, reset_zoom=False):
        if reset_zoom:
            self.zoom_factor = 1.0
            
        selection = self.listbox.curselection()
        if not selection: 
            self._update_preview_canvas(None, "페이지를 선택하면 미리보기가 표시됩니다.")
            return
        idx = selection[0]
        page_info = self.pages[idx]
        
        try:
            doc = page_info['doc']
            page_num = page_info['page_num']
            page = doc.load_page(page_num)
            
            # Apply user-defined rotation on top of original rotation
            total_rot = (page.rotation + page_info['rotation']) % 360
            
            # 줌 팩터 적용 (기본 크기를 1.0으로 설정하고 zoom_factor 곱함)
            scale = 1.0 * getattr(self, 'zoom_factor', 1.0)
            mat = fitz.Matrix(scale, scale).prerotate(page_info['rotation'])
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            mode = "RGBA" if pix.alpha else "RGB"
            img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            
            self.current_image = ImageTk.PhotoImage(img)
            self._update_preview_canvas(self.current_image, "")
        except Exception as e:
            self._update_preview_canvas(None, f"미리보기 오류: {e}")
            
    def on_select(self, event):
        self.update_preview()
        
    def on_mouse_wheel(self, event):
        if not self.pages or not self.listbox.curselection():
            return
            
        # 확대/축소 비율 계산
        if getattr(event, 'num', 0) == 4 or getattr(event, 'delta', 0) > 0:
            self.zoom_factor *= 1.2  # 20% 확대
        elif getattr(event, 'num', 0) == 5 or getattr(event, 'delta', 0) < 0:
            self.zoom_factor /= 1.2  # 20% 축소
            
        # 줌 제한 (너무 작아지거나 커지지 않게)
        self.zoom_factor = max(0.1, min(self.zoom_factor, 10.0))
        self.update_preview()
        
    def rotate_selected(self, angle):
        selection = self.listbox.curselection()
        if not selection: return
        for idx in selection:
            self.pages[idx]['rotation'] = (self.pages[idx]['rotation'] + angle) % 360
        self.update_preview()
            
    def move_selected(self, direction):
        selection = list(self.listbox.curselection())
        if not selection: return
        
        if direction == -1: # Move Up
            if selection[0] == 0: return
            selection.sort()
            for idx in selection:
                text = self.listbox.get(idx)
                self.listbox.delete(idx)
                self.listbox.insert(idx - 1, text)
                self.pages.insert(idx - 1, self.pages.pop(idx))
                self.listbox.selection_set(idx - 1)
        elif direction == 1: # Move Down
            if selection[-1] == len(self.pages) - 1: return
            selection.sort(reverse=True)
            for idx in selection:
                text = self.listbox.get(idx)
                self.listbox.delete(idx)
                self.listbox.insert(idx + 1, text)
                self.pages.insert(idx + 1, self.pages.pop(idx))
                self.listbox.selection_set(idx + 1)
        self.update_preview()
            
    def delete_selected(self):
        selection = list(self.listbox.curselection())
        if not selection:
            messagebox.showwarning("경고", "삭제할 페이지를 선택해주세요.")
            return
        
        selection.sort(reverse=True)
        for idx in selection:
            self.listbox.delete(idx)
            self.pages.pop(idx)
        self.update_preview()
            
    def save_pdf(self):
        if not self.pages:
            messagebox.showwarning("경고", "저장할 페이지가 없습니다.")
            return
            
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not filepath: return
        
        try:
            out_pdf = fitz.open()
            for p in self.pages:
                out_pdf.insert_pdf(p['doc'], from_page=p['page_num'], to_page=p['page_num'])
                # apply rotation to the newly inserted page (which is the last page)
                new_page = out_pdf[-1]
                total_rot = (new_page.rotation + p['rotation']) % 360
                new_page.set_rotation(total_rot)
                
            out_pdf.save(filepath)
            out_pdf.close()
            messagebox.showinfo("성공", f"성공적으로 저장되었습니다:\n{filepath}")
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFEditor(root)
    root.mainloop()
