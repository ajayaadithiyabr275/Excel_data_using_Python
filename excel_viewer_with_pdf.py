import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from fpdf import FPDF

def load_excel():
    global df_original
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if file_path:
        try:
            df_original = pd.read_excel(file_path)
            update_treeview(df_original)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file:\n{e}")

def update_treeview(df):
    tree.delete(*tree.get_children())
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor='center')
    for row in df.to_numpy().tolist():
        tree.insert("", "end", values=row)

def filter_data(*args):
    keyword = search_var.get().lower()
    if keyword == "":
        update_treeview(df_original)
    else:
        try:
            mask = df_original.astype(str).apply(lambda col: col.str.lower().str.contains(keyword), axis=0).any(axis=1)
            filtered_df = df_original[mask]
            update_treeview(filtered_df)
        except Exception as e:
            messagebox.showerror("Search Error", str(e))

def export_to_pdf():
    rows = [tree.item(child)['values'] for child in tree.get_children()]
    if not rows:
        messagebox.showwarning("No Data", "There is no data to export.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not file_path:
        return

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    col_width = 190 / len(tree["columns"])
    row_height = 10

    for col in tree["columns"]:
        pdf.cell(col_width, row_height, str(col), border=1)
    pdf.ln(row_height)

    for row in rows:
        for item in row:
            pdf.cell(col_width, row_height, str(item), border=1)
        pdf.ln(row_height)

    pdf.output(file_path)
    messagebox.showinfo("Success", f"Data exported to:\n{file_path}")

root = tk.Tk()
root.title("Interactive Data Viewer")
root.geometry("1100x600")

top_frame = tk.Frame(root)
top_frame.pack(fill='x', padx=10, pady=5)

search_var = tk.StringVar()
search_var.trace_add("write", filter_data)

tk.Label(top_frame, text="Search:", font=("Arial", 12)).pack(side="left")
search_entry = tk.Entry(top_frame, textvariable=search_var, width=40, font=("Arial", 12))
search_entry.pack(side="left", padx=5)

load_btn = tk.Button(top_frame, text="Load Excel File", font=("Arial", 12), command=load_excel)
load_btn.pack(side="left", padx=5)

export_btn = tk.Button(top_frame, text="Export to PDF", font=("Arial", 12), command=export_to_pdf)
export_btn.pack(side="left", padx=5)

table_frame = tk.Frame(root)
table_frame.pack(fill='both', expand=True, padx=10, pady=10)

tree = ttk.Treeview(table_frame)
vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

tree.grid(row=0, column=0, sticky='nsew')
vsb.grid(row=0, column=1, sticky='ns')
hsb.grid(row=1, column=0, sticky='ew')

table_frame.grid_rowconfigure(0, weight=1)
table_frame.grid_columnconfigure(0, weight=1)

# --- Enhance UI Styling ---
style = ttk.Style()
style.theme_use("clam")  # Options: 'default', 'alt', 'clam', 'vista', 'xpnative'

style.configure("Treeview",
    background="#f0f0f0",
    foreground="black",
    rowheight=25,
    fieldbackground="#e8f1ff",
    font=('Arial', 11)
)
style.configure("Treeview.Heading", 
    font=('Arial', 12, 'bold'),
    background="#1f4e79",
    foreground="white"
)
style.map("Treeview", background=[("selected", "#4a90e2")])


root.mainloop()
