import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image
import os
import threading

selected_file_path = ""

def select_file():
    global selected_file_path
    file_path = filedialog.askopenfilename(
        title="도장/서명 이미지 선택",
        filetypes=(("이미지 파일", "*.png *.jpg *.jpeg *.bmp"), ("모든 파일", "*.*"))
    )
    if file_path:
        selected_file_path = file_path
        lbl_file.config(text=f"선택됨: {os.path.basename(file_path)}")

def generate_images():
    global selected_file_path
    if not selected_file_path:
        messagebox.showwarning("파일 선택", "먼저 굵기를 조절할 그림(서명) 파일을 선택해 주세요!")
        return

    # 저장할 폴더 선택창 띄우기
    save_dir = filedialog.askdirectory(title="변환된 이미지를 저장할 폴더 선택")
    
    if not save_dir:
        return

    btn_select.config(state='disabled')
    btn_generate.config(state='disabled')
    lbl_status.config(text="이미지 변환 및 저장 중... 잠시만 기다려주세요.", foreground="blue")
    root.update()

    try:
        filename = os.path.basename(selected_file_path)
        name_only, _ = os.path.splitext(filename)

        original_img = Image.open(selected_file_path).convert('RGBA')
        
        # 굵기가 일정하게 먹히도록 해상도를 적당히(높이 200px 기준) 통일
        aspect_ratio = original_img.width / original_img.height
        target_height = 200
        target_width = int(target_height * aspect_ratio)
        base_img = original_img.resize((target_width, target_height), Image.Resampling.LANCZOS)
        
        for level in range(1, 5):
            img = base_img.copy()
            
            if level > 1:
                # 굵기 정도 (level 2: 1px, level 3: 2px, level 4: 4px 반경)
                offset = 1 if level == 2 else (2 if level == 3 else 4)
                
                canvas = Image.new('RGBA', img.size, (255, 255, 255, 0))
                for dx in range(-offset, offset + 1):
                    for dy in range(-offset, offset + 1):
                        if dx*dx + dy*dy <= offset*offset:
                            temp = Image.new('RGBA', img.size, (255, 255, 255, 0))
                            temp.paste(img, (dx, dy))
                            canvas = Image.alpha_composite(canvas, temp)
                img = canvas
                
            level_name = "원본" if level == 1 else f"굵기{level-1}단계"
            out_path = os.path.join(save_dir, f"{name_only}_{level_name}.png")
            img.save(out_path, 'PNG')

        lbl_status.config(text=f"완료! 저장 위치: {save_dir}", foreground="green")
        messagebox.showinfo("생성 완료", f"선택하신 폴더에 4단계 굵기의 서명 이미지가 각각 저장되었습니다!\n\n저장 위치: {save_dir}")
        os.startfile(save_dir) # 폴더 열어주기
        
    except Exception as e:
        lbl_status.config(text="오류가 발생했습니다.", foreground="red")
        messagebox.showerror("오류", str(e))
    finally:
        btn_select.config(state='normal')
        btn_generate.config(state='normal')

def on_click_generate():
    threading.Thread(target=generate_images, daemon=True).start()

# --- UI Setup ---
root = tk.Tk()
root.title("서명 굵기 조절 후 폴더 저장기")
root.geometry("450x250")
root.resizable(False, False)

style = ttk.Style(root)
style.theme_use('clam')

frame = ttk.Frame(root, padding=20)
frame.pack(fill='both', expand=True)

lbl_title = ttk.Label(frame, text="서명 굵기 단계별 개별 파일 저장", font=('맑은 고딕', 14, 'bold'))
lbl_title.pack(pady=(0, 10))

lbl_desc = ttk.Label(frame, text="서명/도장 파일을 선택하고 폴더를 지정하면\n해당 폴더에 원본부터 3단계 굵기까지 총 4장의 PNG 이미지가 저장됩니다.", justify="center", font=('맑은 고딕', 10))
lbl_desc.pack(pady=(0, 10))

btn_select = ttk.Button(frame, text="📂 그림 파일 찾아보기...", command=select_file, width=25)
btn_select.pack(pady=5)

lbl_file = ttk.Label(frame, text="선택된 파일: 없음", font=('맑은 고딕', 9), foreground="gray")
lbl_file.pack(pady=(0, 10))

btn_generate = ttk.Button(frame, text="이미지 폴더에 저장하기", command=on_click_generate, width=25)
btn_generate.pack(pady=5)

lbl_status = ttk.Label(frame, text="대기 중...", font=('맑은 고딕', 9))
lbl_status.pack(pady=(5, 0))

root.mainloop()
