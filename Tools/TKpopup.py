import tkinter as tk

def popupmsg(msg, title):
    root = tk.Tk()
    root.title(title)
    root.geometry("500x200")
    label = tk.Label(root, text=msg)
    label.pack(side="top", fill="x", pady=10)
    B1 = tk.Button(root, text="Okay", command = root.destroy)
    B1.pack()
    root.mainloop()

popupmsg('ALERT', 'Pop up Window')
