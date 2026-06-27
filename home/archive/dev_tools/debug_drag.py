import tkinter as tk
from tkinter import ttk

def log_event(e, name):
    print(f"[{name}] {e.type} on {e.widget.winfo_class()} ({e.widget})")

root = tk.Tk()
canvas = tk.Canvas(root, bg='lightgray')
canvas.pack(fill='both', expand=True)
frame = ttk.Frame(canvas)
canvas.create_window((0,0), window=frame, anchor="nw")

e1 = ttk.Entry(frame)
e1.pack(padx=50, pady=50)
e1.insert(0, "Drag me!")

root.bind_all("<B1-Motion>", lambda e: log_event(e, "bind_all B1-Motion"), add="+")
e1.bind("<B1-Motion>", lambda e: log_event(e, "e1 B1-Motion"), add="+")

root.after(3000, root.destroy)
root.mainloop()
