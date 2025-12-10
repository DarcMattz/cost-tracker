import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import os
from datetime import date

EXCEL_FILE = "cost_data.xlsx"

# Ensure Excel file exists
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CostData"
    ws.append(["Item Description","Category","Unit","Material","Labor","Other","Total","Brand","Date","Remarks"])
    wb.save(EXCEL_FILE)

# Load data
def load_data():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        data.append(row)
    return data

# Save data
def save_data(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CostData"

    # Write headers
    ws.append(["Item Description","Category","Unit","Material","Labor","Other","Total","Brand","Date","Remarks"])

    # Write data rows
    for row in data:
        ws.append(list(row))  # ensure tuple becomes list
    
    wb.save(EXCEL_FILE)

# Add/Edit Item
def save_item():
    global data, edit_index

    item_desc = item_entry.get().strip()
    category = category_entry.get().strip()
    unit = unit_entry.get().strip()
    brand = brand_entry.get().strip()

    try:
        material_cost = float(material_entry.get() or 0)
        labor_cost = float(labor_entry.get() or 0)
        other_cost = float(other_entry.get() or 0)
    except ValueError:
        messagebox.showerror("Error", "Cost values must be numeric.")
        return

    item_date = date_entry.get().strip() or str(date.today())
    notes = notes_entry.get().strip()
    total = material_cost + labor_cost + other_cost

    if not item_desc:
        messagebox.showwarning("Input Error", "Item Description is required.")
        return

    new_row = (item_desc, category, unit, material_cost, labor_cost, other_cost, total, brand, item_date, notes)

    if edit_index is not None:
        data[edit_index] = new_row
        edit_index = None
    else:
        data.append(new_row)

    save_data(data)
    render_table()
    clear_form()

# Delete Item
def delete_item():
    global data
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Select", "Select an item to delete.")
        return
    idx = int(selected[0])
    if messagebox.askyesno("Confirm", "Delete this item?"):
        data.pop(idx)
        save_data(data)
        render_table()

# Edit Item
def edit_item():
    global edit_index
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Select", "Select an item to edit.")
        return

    idx = int(selected[0])
    row = data[idx]
    edit_index = idx

    item_entry.delete(0, tk.END)
    item_entry.insert(0, row[0])

    category_entry.delete(0, tk.END)
    category_entry.insert(0, row[1])

    unit_entry.delete(0, tk.END)
    unit_entry.insert(0, row[2])

    material_entry.delete(0, tk.END)
    material_entry.insert(0, row[3])

    labor_entry.delete(0, tk.END)
    labor_entry.insert(0, row[4])

    other_entry.delete(0, tk.END)
    other_entry.insert(0, row[5])

    brand_entry.delete(0, tk.END)
    brand_entry.insert(0, row[7])

    date_entry.delete(0, tk.END)
    date_entry.insert(0, row[8])

    notes_entry.delete(0, tk.END)
    notes_entry.insert(0, row[9])

# Clear form
def clear_form():
    global edit_index
    item_entry.delete(0, tk.END)
    category_entry.delete(0, tk.END)
    unit_entry.delete(0, tk.END)
    material_entry.delete(0, tk.END)
    labor_entry.delete(0, tk.END)
    other_entry.delete(0, tk.END)
    brand_entry.delete(0, tk.END)
    date_entry.delete(0, tk.END)
    date_entry.insert(0, str(date.today()))
    notes_entry.delete(0, tk.END)
    edit_index = None

# Render Table (with search)
def render_table():
    query = search_var.get().lower()
    for i in tree.get_children():
        tree.delete(i)

    for idx, row in enumerate(data):
        if query in str(row[0]).lower() or query in str(row[1]).lower():
            tree.insert("", "end", iid=idx, values=row)

# --- GUI ---
root = tk.Tk()
root.title("Cost Database")
root.configure(bg="#f0f2f5")

try:
    root.state("zoomed")
except:
    root.geometry("1200x700")

# Search Frame
search_frame = tk.Frame(root, bg="#f0f2f5")
search_frame.pack(fill=tk.X, padx=10, pady=5)

search_var = tk.StringVar()
tk.Label(search_frame, text="Search:", font=("Arial", 12), bg="#f0f2f5").pack(side=tk.LEFT)
search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Arial", 12), width=40)
search_entry.pack(side=tk.LEFT, padx=10)
search_entry.bind("<KeyRelease>", lambda e: render_table())

# Table Frame
table_frame = tk.Frame(root)
table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

columns = ("Item Description","Category","Unit","Material","Labor","Other","Total","Brand","Date","Remarks")
tree = ttk.Treeview(table_frame, columns=columns, show="headings")

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=150 if col in ["Item Description","Remarks","Brand"] else 110)

tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Form Frame
form_frame = tk.LabelFrame(root, text="Add / Edit Item", padx=10, pady=10)
form_frame.pack(fill=tk.X, padx=10, pady=10)

tk.Label(form_frame, text="Item Description").grid(row=0, column=0)
item_entry = tk.Entry(form_frame)
item_entry.grid(row=0, column=1)

tk.Label(form_frame, text="Category").grid(row=0, column=2)
category_entry = tk.Entry(form_frame)
category_entry.grid(row=0, column=3)

tk.Label(form_frame, text="Unit").grid(row=1, column=0)
unit_entry = tk.Entry(form_frame)
unit_entry.grid(row=1, column=1)

tk.Label(form_frame, text="Material").grid(row=1, column=2)
material_entry = tk.Entry(form_frame)
material_entry.grid(row=1, column=3)

tk.Label(form_frame, text="Labor").grid(row=2, column=0)
labor_entry = tk.Entry(form_frame)
labor_entry.grid(row=2, column=1)

tk.Label(form_frame, text="Other").grid(row=2, column=2)
other_entry = tk.Entry(form_frame)
other_entry.grid(row=2, column=3)

tk.Label(form_frame, text="Brand").grid(row=3, column=0)
brand_entry = tk.Entry(form_frame)
brand_entry.grid(row=3, column=1)

tk.Label(form_frame, text="Date").grid(row=3, column=2)
date_entry = tk.Entry(form_frame)
date_entry.grid(row=3, column=3)
date_entry.insert(0, str(date.today()))

tk.Label(form_frame, text="Remarks").grid(row=4, column=0)
notes_entry = tk.Entry(form_frame, width=40)
notes_entry.grid(row=4, column=1, columnspan=3)

# Buttons
btn_frame = tk.Frame(form_frame)
btn_frame.grid(row=5, column=0, columnspan=4, pady=10)

tk.Button(btn_frame, text="Save", bg="#4CAF50", fg="white", width=12, command=save_item).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Clear", bg="#f0ad4e", fg="white", width=12, command=clear_form).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Edit Selected", bg="#2196F3", fg="white", width=12, command=edit_item).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Delete Selected", bg="#f44336", fg="white", width=12, command=delete_item).pack(side=tk.LEFT, padx=5)

# Load & render
data = load_data()
edit_index = None
render_table()

# Watermark
watermark = tk.Label(root, text="Jibee", font=("Arial", 10, "italic"),
                     bg="#f0f2f5", fg="#999999")
watermark.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)


root.mainloop()
