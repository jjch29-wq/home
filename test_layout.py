import tkinter as tk
from tkinter import ttk
root = tk.Tk()
root.geometry('800x800')

c = tk.Canvas(root, bg='green')
c.pack(fill='both', expand=True)

f = ttk.Frame(c)
cw = c.create_window((0,0), window=f, anchor='nw')

def on_conf(e):
    req_h = f.winfo_reqheight()
    target_h = max(req_h, e.height)
    print("Canvas height:", e.height, "Req:", req_h, "Target:", target_h)
    c.itemconfig(cw, width=e.width, height=target_h)

c.bind('<Configure>', on_conf)

# mimic setup
f.grid_rowconfigure(1, weight=1)
f.grid_columnconfigure(0, weight=1)

top = ttk.LabelFrame(f, text='Top Form')
top.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
ttk.Label(top, text='Height=350').pack(pady=175)

bot = ttk.PanedWindow(f, orient='horizontal')
bot.grid(row=1, column=0, sticky='nsew', padx=5, pady=10)

left = ttk.LabelFrame(bot, text='Left Pane')
bot.add(left, weight=1)
tv = ttk.Treeview(left, height=9)
tv.pack(side='left', fill='both', expand=True)

root.update()
print("f height:", f.winfo_height())
print("bot height:", bot.winfo_height())
print("left height:", left.winfo_height())
print("tv height:", tv.winfo_height())
