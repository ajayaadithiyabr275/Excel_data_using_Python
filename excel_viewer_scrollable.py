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
root.title("Excel Viewer with Scrollbars")
root.geometry("800x500")

frame = tk.Frame(root)
frame.pack(pady=10)

btn = tk.Button(frame, text="Load Excel File", command=load_excel)
btn.pack()

tree_frame = tk.Frame(root)
tree_frame.pack(expand=True, fill='both')

# Add scrollbars
tree_scroll_y = tk.Scrollbar(tree_frame, orient="vertical")
tree_scroll_y.pack(side='right', fill='y')

tree_scroll_x = tk.Scrollbar(tree_frame, orient="horizontal")
tree_scroll_x.pack(side='bottom', fill='x')

tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
tree.pack(expand=True, fill='both')

tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)

root.mainloop()
