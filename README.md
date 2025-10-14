```
import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd

def browse_file(entry_widget):
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

def submit_action():
    file1 = file1_entry.get()
    file2 = file2_entry.get()
    selected_date = date_entry.get_date()

    if not file1 or not file2:
        messagebox.showerror("Error", "Please select both Excel files.")
        return

    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
        messagebox.showinfo(
            "Files Loaded",
            f"✅ Both Excel files loaded successfully!\n\n"
            f"📅 Selected Date: {selected_date}\n\n"
            f"File 1: {len(df1)} rows\n"
            f"File 2: {len(df2)} rows"
        )
        root.destroy()  # Close the UI after success
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel files:\n\n{e}")

# --- Main Window ---
root = tk.Tk()
root.title("Excel File and Date Selector")
root.geometry("500x280")
root.resizable(False, False)

# File 1
tk.Label(root, text="Select First Excel File:").grid(row=0, column=0, padx=10, pady=20, sticky="w")
file1_entry = tk.Entry(root, width=45)
file1_entry.grid(row=0, column=1, padx=5)
tk.Button(root, text="Browse", command=lambda: browse_file(file1_entry)).grid(row=0, column=2, padx=5)

# File 2
tk.Label(root, text="Select Second Excel File:").grid(row=1, column=0, padx=10, pady=20, sticky="w")
file2_entry = tk.Entry(root, width=45)
file2_entry.grid(row=1, column=1, padx=5)
tk.Button(root, text="Browse", command=lambda: browse_file(file2_entry)).grid(row=1, column=2, padx=5)

# Date picker
tk.Label(root, text="Select Date:").grid(row=2, column=0, padx=10, pady=20, sticky="w")
date_entry = DateEntry(root, width=18, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
date_entry.grid(row=2, column=1, padx=5, sticky="w")

# Submit button
tk.Button(root, text="Load Files", width=20, bg="#4CAF50", fg="white", command=submit_action).grid(row=3, column=1, pady=30)

root.mainloop()
