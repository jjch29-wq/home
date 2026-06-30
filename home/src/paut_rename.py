import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

class PAUTRenameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PAUT 데이터 파일 일괄 이름 변경기")
        self.root.geometry("600x500")
        
        # --- Variables ---
        self.dir_var = tk.StringVar()
        self.find_var = tk.StringVar()
        self.replace_var = tk.StringVar()
        self.file_list = [] # (original_path, new_name)
        
        self.create_widgets()
        
    def create_widgets(self):
        # 1. Directory Selection
        df = tk.LabelFrame(self.root, text="1. 폴더 선택 (PAUT 데이터 위치)")
        df.pack(fill="x", padx=10, pady=5)
        tk.Entry(df, textvariable=self.dir_var, state="readonly").pack(side="left", fill="x", expand=True, padx=5, pady=5)
        tk.Button(df, text="폴더 열기", command=self.browse_dir).pack(side="right", padx=5, pady=5)
        
        # 2. Find and Replace
        rf = tk.LabelFrame(self.root, text="2. 이름 변경 규칙")
        rf.pack(fill="x", padx=10, pady=5)
        
        tk.Label(rf, text="찾을 단어:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.find_entry = tk.Entry(rf, textvariable=self.find_var, width=20)
        self.find_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.find_var.trace_add("write", self.update_preview)
        
        tk.Label(rf, text="바꿀 단어:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.replace_entry = tk.Entry(rf, textvariable=self.replace_var, width=20)
        self.replace_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        self.replace_var.trace_add("write", self.update_preview)
        
        tk.Button(rf, text="적용 미리보기", command=self.update_preview, bg="lightblue").grid(row=0, column=4, padx=10, pady=5)

        # 3. Preview Treeview
        pf = tk.LabelFrame(self.root, text="3. 변경 미리보기 (PAUT 데이터 및 그림 파일 표시)")
        pf.pack(fill="both", expand=True, padx=10, pady=5)
        
        columns = ("Original", "New")
        self.tree = ttk.Treeview(pf, columns=columns, show="headings")
        self.tree.heading("Original", text="원래 이름")
        self.tree.heading("New", text="바뀔 이름")
        self.tree.column("Original", width=250)
        self.tree.column("New", width=250)
        
        scrollbar = ttk.Scrollbar(pf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 4. Action Buttons
        bf = tk.Frame(self.root)
        bf.pack(fill="x", padx=10, pady=10)
        tk.Button(bf, text="실제 이름 변경 실행", font=("Arial", 12, "bold"), bg="salmon", command=self.execute_rename).pack(fill="x")
        
    def browse_dir(self):
        folder = filedialog.askdirectory(title="PAUT 파일이 있는 폴더 선택")
        if folder:
            self.dir_var.set(folder)
            self.update_preview()
            
    def update_preview(self, *args):
        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        self.file_list = []
        folder = self.dir_var.get()
        if not folder or not os.path.isdir(folder):
            return
            
        find_text = self.find_var.get()
        replace_text = self.replace_var.get()
        
        # Scan for .opd, .nde and image files
        p = Path(folder)
        all_files = list(p.glob("*.opd")) + list(p.glob("*.nde"))
        for ext in ["*.png", "*.jpg", "*.jpeg", "*.bmp", "*.PNG", "*.JPG", "*.JPEG", "*.BMP"]:
            all_files.extend(list(p.glob(ext)))
        
        for file_path in all_files:
            original_name = file_path.name
            
            if find_text:
                new_name = original_name.replace(find_text, replace_text)
            else:
                new_name = original_name
                
            self.file_list.append((file_path, new_name))
            
            # Insert into tree
            # If changed, show in blue
            item_id = self.tree.insert("", "end", values=(original_name, new_name))
            if original_name != new_name:
                self.tree.item(item_id, tags=("changed",))
                
        self.tree.tag_configure("changed", foreground="blue")
        
    def execute_rename(self):
        if not self.file_list:
            messagebox.showwarning("경고", "변경할 파일이 없습니다.")
            return
            
        find_text = self.find_var.get()
        if not find_text:
            messagebox.showwarning("경고", "찾을 단어를 입력해주세요.")
            return
            
        # Count how many will change
        to_change = [f for f in self.file_list if f[0].name != f[1]]
        if not to_change:
            messagebox.showinfo("안내", "이름이 바뀔 파일이 없습니다.")
            return
            
        confirm = messagebox.askyesno("확인", f"총 {len(to_change)}개의 파일 이름을 변경하시겠습니까?\n이 작업은 되돌릴 수 없습니다!")
        if not confirm:
            return
            
        success_count = 0
        for file_path, new_name in to_change:
            new_path = file_path.parent / new_name
            try:
                os.rename(file_path, new_path)
                success_count += 1
            except Exception as e:
                print(f"Error renaming {file_path.name}: {e}")
                
        messagebox.showinfo("완료", f"총 {success_count}개의 파일 이름이 성공적으로 변경되었습니다.")
        self.update_preview()

if __name__ == "__main__":
    root = tk.Tk()
    app = PAUTRenameApp(root)
    root.mainloop()
