import tkinter as tk
root = tk.Tk()
root.title("테스트 창")
root.geometry("300x200+300+300")
root.configure(bg='red')
lbl = tk.Label(root, text="이 창이 보이면 정상!", font=("Arial", 16), bg='red', fg='white')
lbl.pack(expand=True)
root.after(5000, root.destroy)
root.mainloop()
