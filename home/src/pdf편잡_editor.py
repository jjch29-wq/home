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
        self.zoom_level = 1.0
        
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
        ttk.Button(ctrl_frame, text="선택 삭제", command=self.delete_selected).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        
        ttk.Button(left_frame, text="저장하기", command=self.save_pdf).pack(fill=tk.X, pady=5)
        
        # Right frame: Preview and Rotation
        right_frame = ttk.Frame(self.root, padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        rot_frame = ttk.Frame(right_frame)
        rot_frame.pack(fill=tk.X, pady=5)
        ttk.Button(rot_frame, text="↺ 왼쪽으로 회전", command=lambda: self.rotate_selected(-90)).pack(side=tk.LEFT, padx=5)
        ttk.Button(rot_frame, text="↻ 오른쪽으로 회전", command=lambda: self.rotate_selected(90)).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(rot_frame, text="🔍 줌 인(+)", command=self.zoom_in).pack(side=tk.RIGHT, padx=5)
        ttk.Button(rot_frame, text="🔍 줌 아웃(-)", command=self.zoom_out).pack(side=tk.RIGHT, padx=5)
        self.zoom_label = ttk.Label(rot_frame, text="100%")
        self.zoom_label.pack(side=tk.RIGHT, padx=5)
        
        self.canvas_frame = ttk.Frame(right_frame)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(self.canvas_frame, bg='white')
        self.scrollbar_y = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar_x = ttk.Scrollbar(self.canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)
        
        self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.preview_image_id = self.canvas.create_image(0, 0, anchor=tk.CENTER)
        self.preview_text_id = self.canvas.create_text(200, 200, text="페이지를 선택하면 미리보기가 표시됩니다.", font=("Arial", 12), fill="gray")
        
        self.canvas.bind('<Configure>', self._center_image)
        self.canvas.bind('<MouseWheel>', self.on_mousewheel_scroll)
        self.canvas.bind('<Control-MouseWheel>', self.on_mousewheel_zoom)
        
        self.current_image = None
        
    def _center_image(self, event=None):
        c_width = self.canvas.winfo_width()
        c_height = self.canvas.winfo_height()
        
        if self.current_image:
            i_width = self.current_image.width()
            i_height = self.current_image.height()
            
            x = max(c_width // 2, i_width // 2)
            y = max(c_height // 2, i_height // 2)
            self.canvas.coords(self.preview_image_id, x, y)
        else:
            self.canvas.coords(self.preview_image_id, c_width // 2, c_height // 2)
            
        self.canvas.coords(self.preview_text_id, c_width // 2, c_height // 2)
        self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
        
    def on_mousewheel_scroll(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
    def on_mousewheel_zoom(self, event):
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()

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
        self.canvas.itemconfig(self.preview_image_id, image='')
        self.canvas.itemconfig(self.preview_text_id, text="페이지를 선택하면 미리보기가 표시됩니다.")
        for filepath in filepaths:
            self.add_document(filepath)
        
    def add_pdf(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if not filepaths: return
        for filepath in filepaths:
            self.add_document(filepath)
        
    def add_document(self, filepath):
        try:
            # Memory load to release file lock immediately
            with open(filepath, "rb") as f:
                pdf_bytes = f.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            self.docs.append(doc)
            
            filename = os.path.basename(filepath)
            for i in range(len(doc)):
                label = f"{filename} - {i+1}쪽"
                self.pages.append({'doc': doc, 'page_num': i, 'label': label, 'rotation': 0})
                self.listbox.insert(tk.END, label)
        except Exception as e:
            messagebox.showerror("오류", f"PDF를 열 수 없습니다:\n{e}")
            
    def update_preview(self):
        selection = self.listbox.curselection()
        if not selection: 
            self.canvas.itemconfig(self.preview_image_id, image='')
            self.canvas.itemconfig(self.preview_text_id, text="페이지를 선택하면 미리보기가 표시됩니다.")
            return
        idx = selection[0]
        page_info = self.pages[idx]
        
        try:
            doc = page_info['doc']
            page_num = page_info['page_num']
            page = doc.load_page(page_num)
            
            # Apply user-defined rotation on top of original rotation
            total_rot = (page.rotation + page_info['rotation']) % 360
            
            # High resolution for better preview
            mat = fitz.Matrix(3.0, 3.0).prerotate(page_info['rotation'])
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            mode = "RGBA" if pix.alpha else "RGB"
            img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            
            # Resize for preview with zoom
            max_w, max_h = int(1200 * self.zoom_level), int(1600 * self.zoom_level)
            img.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
            
            self.current_image = ImageTk.PhotoImage(img)
            self.canvas.itemconfig(self.preview_image_id, image=self.current_image)
            self.canvas.itemconfig(self.preview_text_id, text="")
            self._center_image()
        except Exception as e:
            self.canvas.itemconfig(self.preview_image_id, image='')
            self.canvas.itemconfig(self.preview_text_id, text=f"미리보기 오류: {e}")
            
    def on_select(self, event):
        self.update_preview()
        
    def rotate_selected(self, angle):
        selection = self.listbox.curselection()
        if not selection: return
        for idx in selection:
            self.pages[idx]['rotation'] = (self.pages[idx]['rotation'] + angle) % 360
        self.update_preview()
            
    def zoom_in(self):
        self.zoom_level += 0.2
        self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
        self.update_preview()
        
    def zoom_out(self):
        self.zoom_level = max(0.2, self.zoom_level - 0.2)
        self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
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
