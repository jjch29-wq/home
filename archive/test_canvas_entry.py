import tkinter as tk
from tkinter import ttk

root = tk.Tk()
canvas = tk.Canvas(root, bg="yellow")
scrollable_frame = tk.Frame(canvas, bg="lightblue")
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.pack(fill="both", expand=True)

nb = ttk.Notebook(scrollable_frame)
nb.pack()
f1 = ttk.Frame(nb)
nb.add(f1, text="Tab 1")

e1 = ttk.Entry(f1)
e1.pack(pady=20)
e1.insert(0, "Try to drag me inside Tab 1")

root.after(5000, lambda: print("Auto-closing"))
root.after(5500, root.destroy)
root.mainloop()
