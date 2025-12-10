# 📘 **Cost Database**

A simple and user-friendly **Cost Database Application** built with **Python**, **Tkinter**, and **OpenPyXL**.
This tool allows you to record, edit, delete, and search cost-related items such as materials, labor, and other expenses.
All records are automatically stored in an Excel file (`cost_data.xlsx`).

---

## 🚀 **Features**

### ✅ **User-Friendly Interface**

- Clean and simple UI designed with Tkinter
- Auto-maximized window
- Organized form panel for adding and editing items
- Bottom-right watermark showing **"Jibee"**

### 📁 **Excel-Based Storage**

- Automatically creates `cost_data.xlsx` if it doesn’t exist
- Saves all entries into an Excel sheet with clear column headers
- Updates the file whenever you add, edit, or delete items

### 🔎 **Search Function**

- Real-time search bar
- Instantly filters results based on Item Description or Category

### 📝 **Item Fields**

Each record can include:

- Item Description
- Category
- Unit
- Material Cost
- Labor Cost
- Other Cost
- Total (auto-calculated)
- Brand
- Date
- Remarks

### ✏ **Record Management**

- Add new items
- Edit selected items
- Delete selected items
- Clear form instantly

---

## 🛠 **Requirements**

Make sure you have:

- **Python 3.8+** (recommended 3.10 or 3.11)
- Pip installed

Required Python packages:

```
pip install openpyxl
```

Tkinter is already included with most Python installations.

---

## ▶ **How to Run the App**

1. Download or clone the project.
2. Open a terminal inside the project folder.
3. Run the app:

```
python index.py
```

The app will automatically:

- Create `cost_data.xlsx` if it doesn't exist
- Load existing data
- Open a full-screen window ready for use

---

## 🧩 **File Structure**

```
project/
│
├── index.py          # Main application
├── cost_data.xlsx    # Auto-created Excel file (if not present)
└── README.md         # This documentation
```

---

## ❗ Known Behavior / Notes

- On first launch, the Excel file is created automatically with headers.
- The app calculates the **Total Cost** automatically.
- The search bar filters results as you type.

---

## 📌 **Future Improvements (Optional Ideas)**

- Export to PDF
- Color themes or dark mode
- Category dropdown options
- Import existing Excel files
- Pagination for large datasets

---

## 👤 **Author**

**Jibee**
Created for personal use and cost tracking.
