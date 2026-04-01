import tkinter as tk
from tkinter import ttk

def log_event(e, name):
    print(f"[{name}] cursor={e.widget.index(tk.INSERT)}, sel={e.widget.select_present()}")

root = tk.Tk()
style = ttk.Style()
style.theme_use('clam')

canvas = tk.Canvas(root, bg='lightgray')
canvas.pack(fill='both', expand=True)
frame = ttk.Frame(canvas)
canvas.create_window((0,0), window=frame, anchor="nw")

e1 = ttk.Entry(frame)
e1.pack(padx=50, pady=50)
e1.insert(0, "Drag me!")
e1.bind("<B1-Motion>", lambda e: log_event(e, "Motion"), add="+")
e1.bind("<ButtonRelease-1>", lambda e: log_event(e, "Release"), add="+")

def sim_drag():
    e1.focus_set()
    e1.event_generate("<Button-1>", x=5, y=10)
    for i in range(5, 100, 5):
        e1.event_generate("<B1-Motion>", x=i, y=10)
    e1.event_generate("<ButtonRelease-1>", x=100, y=10)
    print("Simulated. Has selection:", e1.select_present())
    root.destroy()

root.after(1000, sim_drag)
root.mainloop()
