import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import os
import subprocess
import pyperclip
from datetime import date

EXCEL_FILE = "cost_data.xlsx"

# Ensure Excel file exists with new column structure
HEADERS = [
    "Item Description","Category","Unit","Material","Total",
    "Brand","Date","Logged By","Reference","Remarks"
]

if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CostData"
    ws.append(HEADERS)
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

    ws.append(HEADERS)
    for row in data:
        ws.append(list(row))

    wb.save(EXCEL_FILE)


# Add/Edit item
def save_item():
    global data, edit_index

    item_desc = item_entry.get().strip()
    category = category_entry.get().strip()
    unit = unit_entry.get().strip()
    brand = brand_entry.get().strip()

    try:
        material_cost = float(material_entry.get() or 0)
    except ValueError:
        messagebox.showerror("Error", "Material cost must be numeric.")
        return

    total = material_cost  # Only material now

    item_date = date_entry.get().strip() or str(date.today())
    logged_by = logged_entry.get().strip()
    reference_path = reference_var.get().strip()
    remarks = notes_entry.get().strip()

    if not item_desc:
        messagebox.showwarning("Input Error", "Item Description is required.")
        return

    new_row = (
        item_desc, category, unit, material_cost, total,
        brand, item_date, logged_by, reference_path, remarks
    )

    if edit_index is not None:
        data[edit_index] = new_row
        edit_index = None
    else:
        data.append(new_row)

    save_data(data)
    render_table()
    clear_form()


# Delete
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


# Edit
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

    brand_entry.delete(0, tk.END)
    brand_entry.insert(0, row[5])

    date_entry.delete(0, tk.END)
    date_entry.insert(0, row[6])

    logged_entry.delete(0, tk.END)
    logged_entry.insert(0, row[7])

    reference_var.set(row[8] or "")

    notes_entry.delete(0, tk.END)
    notes_entry.insert(0, row[9])


# Clear
def clear_form():
    global edit_index
    item_entry.delete(0, tk.END)
    category_entry.delete(0, tk.END)
    unit_entry.delete(0, tk.END)
    material_entry.delete(0, tk.END)
    brand_entry.delete(0, tk.END)
    date_entry.delete(0, tk.END)
    date_entry.insert(0, str(date.today()))
    logged_entry.delete(0, tk.END)
    reference_var.set("")
    notes_entry.delete(0, tk.END)
    edit_index = None


# Browse reference file
def browse_reference():
    filepath = filedialog.askopenfilename()
    if filepath:
        reference_var.set(filepath)


# Right-Click Menu Actions
def open_reference():
    selected = tree.selection()
    if not selected:
        return

    idx = int(selected[0])
    ref = data[idx][8]

    if not ref:
        messagebox.showinfo("No Reference", "No reference assigned.")
        return

    if not os.path.exists(ref):
        messagebox.showerror("Error", "Referenced file does not exist.")
        return

    try:
        os.startfile(ref)
    except Exception as e:
        messagebox.showerror("Error", str(e))


def copy_reference():
    selected = tree.selection()
    if not selected:
        return
    idx = int(selected[0])
    ref = data[idx][8] or ""
    pyperclip.copy(ref)


def paste_reference():
    global data
    selected = tree.selection()
    if not selected:
        return

    idx = int(selected[0])
    paste_val = pyperclip.paste()

    row = list(data[idx])
    row[8] = paste_val
    data[idx] = tuple(row)

    save_data(data)
    render_table()


# Right-click popup
def show_context_menu(event):
    selected = tree.identify_row(event.y)
    if selected:
        tree.selection_set(selected)
        menu.post(event.x_root, event.y_root)


# Render table
def render_table():
    query = search_var.get().lower()
    for i in tree.get_children():
        tree.delete(i)

    for idx, row in enumerate(data):
        if query in str(row[0]).lower() or query in str(row[1]).lower():
            tree.insert("", "end", iid=idx, values=row)


# GUI Setup
root = tk.Tk()
root.title("Cost Database")
root.configure(bg="#f0f2f5")

try:
    root.state("zoomed")
except:
    root.geometry("1300x750")

# Search
search_frame = tk.Frame(root, bg="#f0f2f5")
search_frame.pack(fill=tk.X, padx=10, pady=5)

search_var = tk.StringVar()
tk.Label(search_frame, text="Search:", font=("Arial", 12), bg="#f0f2f5").pack(side=tk.LEFT)
search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Arial", 12), width=40)
search_entry.pack(side=tk.LEFT, padx=10)
search_entry.bind("<KeyRelease>", lambda e: render_table())

# Table
table_frame = tk.Frame(root)
table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

columns = HEADERS
tree = ttk.Treeview(table_frame, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=150)

tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Context menu
menu = tk.Menu(root, tearoff=0)
menu.add_command(label="Open Reference", command=open_reference)
menu.add_command(label="Copy Reference", command=copy_reference)
menu.add_command(label="Paste Reference", command=paste_reference)

tree.bind("<Button-3>", show_context_menu)

# Form
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

tk.Label(form_frame, text="Brand").grid(row=2, column=0)
brand_entry = tk.Entry(form_frame)
brand_entry.grid(row=2, column=1)

tk.Label(form_frame, text="Date").grid(row=2, column=2)
date_entry = tk.Entry(form_frame)
date_entry.grid(row=2, column=3)
date_entry.insert(0, str(date.today()))

tk.Label(form_frame, text="Logged By").grid(row=3, column=0)
logged_entry = tk.Entry(form_frame)
logged_entry.grid(row=3, column=1)

tk.Label(form_frame, text="Reference").grid(row=3, column=2)
reference_var = tk.StringVar()
reference_entry = tk.Entry(form_frame, textvariable=reference_var, width=30)
reference_entry.grid(row=3, column=3)
tk.Button(form_frame, text="Browse", command=browse_reference).grid(row=3, column=4, padx=5)

tk.Label(form_frame, text="Remarks").grid(row=4, column=0)
notes_entry = tk.Entry(form_frame, width=40)
notes_entry.grid(row=4, column=1, columnspan=3)

btn_frame = tk.Frame(form_frame)
btn_frame.grid(row=5, column=0, columnspan=4, pady=10)

tk.Button(btn_frame, text="Save", bg="#4CAF50", fg="white", width=12, command=save_item).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Clear", bg="#f0ad4e", fg="white", width=12, command=clear_form).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Edit Selected", bg="#2196F3", fg="white", width=12, command=edit_item).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Delete Selected", bg="#f44336", fg="white", width=12, command=delete_item).pack(side=tk.LEFT, padx=5)

# Watermark
watermark = tk.Label(root, text="Jibee", font=("Arial", 10, "italic"),
                     bg="#f0f2f5", fg="#999999")
watermark.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)


# Load & render
data = load_data()
edit_index = None
render_table()

root.mainloop()
