import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("드래그 환경 검증 테스트")
root.geometry("400x300")

style = ttk.Style()
style.theme_use('clam')

canvas = tk.Canvas(root, bg="#f0f0f0")
canvas.pack(fill="both", expand=True, pady=10)
frame = tk.Frame(canvas)
canvas.create_window((0,0), window=frame, anchor="nw")

nb = ttk.Notebook(frame)
nb.pack(fill='both', expand=True)

tab1 = ttk.Frame(nb)
nb.add(tab1, text="테스트 탭")

ttk.Label(tab1, text="이 항목이 드래그가 안되는지 테스트해 주세요 (clam + exportselection=False)").pack()

e1 = ttk.Entry(tab1, exportselection=False)
e1.pack(pady=10, fill='x', padx=20)
e1.insert(0, "복잡한 환경: 이 텍스트를 드래그 해 보세요.")

root.mainloop()
