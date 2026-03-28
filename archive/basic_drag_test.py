import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("드래그 테스트 (정상 동작 확인용)")
root.geometry("400x300")

ttk.Label(root, text="아래 칸에서 텍스트를 드래그해 보세요.").pack(pady=10)

e1 = ttk.Entry(root)
e1.pack(pady=10, fill='x', padx=20)
e1.insert(0, "여기를 드래그해서 파란색 블록이 생기는지 확인하세요.")

e2 = tk.Entry(root)
e2.pack(pady=10, fill='x', padx=20)
e2.insert(0, "이 칸도 드래그가 되는지 확인해 보세요.")

# 캔버스 내부 테스트
canvas = tk.Canvas(root, bg="#f0f0f0", height=100)
canvas.pack(fill="both", expand=True, pady=10)
frame = ttk.Frame(canvas)
canvas.create_window((0,0), window=frame, anchor="nw")

e3 = ttk.Entry(frame)
e3.pack(pady=10, fill='x', padx=20)
e3.insert(0, "캔버스 내부의 칸입니다. 드래그 해보세요.")

root.mainloop()
