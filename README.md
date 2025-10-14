import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry

def browse_file(entry_widget):
    """Open file dialog and update the corresponding entry box"""
    file_path = filedialog.askopenfilename(title="Select a file")
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

def submit_action():
    """Fetch and show selected file paths and date"""
    file1 = file1_entry.get()
    file2 = file2_entry.get()
    selected_date = date_entry.get_date()

    if not file1 or not file2:
        messagebox.showerror("Error", "Please select both files.")
        return

    messagebox.showinfo(
        "Selection Summary",
        f"File 1: {file1}\n"
        f"File 2: {file2}\n"
        f"Selected Date: {selected_date}"
    )

# Create main window
root = tk.Tk()
root.title("File and Date Selector")
root.geometry("450x250")
root.resizable(False, False)

# --- File 1 selection ---
tk.Label(root, text="Select First File:").grid(row=0, column=0, padx=10, pady=15, sticky="w")
file1_entry = tk.Entry(root, width=40)
file1_entry.grid(row=0, column=1, padx=5)
tk.Button(root, text="Browse", command=lambda: browse_file(file1_entry)).grid(row=0, column=2, padx=5)

# --- File 2 selection ---
tk.Label(root, text="Select Second File:").grid(row=1, column=0, padx=10, pady=15, sticky="w")
file2_entry = tk.Entry(root, width=40)
file2_entry.grid(row=1, column=1, padx=5)
tk.Button(root, text="Browse", command=lambda: browse_file(file2_entry)).grid(row=1, column=2, padx=5)

# --- Date selection ---
tk.Label(root, text="Select Date:").grid(row=2, column=0, padx=10, pady=15, sticky="w")
date_entry = DateEntry(root, width=15, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
date_entry.grid(row=2, column=1, padx=5, sticky="w")

# --- Submit button ---
tk.Button(root, text="Submit", width=15, command=submit_action, bg="#4CAF50", fg="white").grid(row=3, column=1, pady=25)

root.mainloop()
