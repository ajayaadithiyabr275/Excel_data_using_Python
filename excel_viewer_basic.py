import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd

def load_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if file_path:
        df = pd.read_excel(file_path)
        tree.delete(*tree.get_children())
        tree["column"] = list(df.columns)
        tree["show"] = "headings"
        for col in df.columns:
            tree.heading(col, text=col)
        for row in df.to_numpy().tolist():
            tree.insert("", "end", values=row)

root = tk.Tk()
root.title("Excel Viewer")

frame = tk.Frame(root)
frame.pack(pady=10)

btn = tk.Button(frame, text="Load Excel File", command=load_excel)
btn.pack()

tree = ttk.Treeview(root)
tree.pack(expand=True, fill='both')

root.mainloop()
