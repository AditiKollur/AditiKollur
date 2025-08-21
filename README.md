```
import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry

# Global variables (to be populated after GUI is closed)
input_excel = None
cust_horis = None
hmo_cust = None
output_folder = None
selected_date = None


def run_gui():
    global input_excel, cust_horis, hmo_cust, output_folder, selected_date

    def select_input_excel():
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        if file:
            entry_input_excel.delete(0, tk.END)
            entry_input_excel.insert(0, file)

    def select_cust_horis():
        file = filedialog.askopenfilename(filetypes=[("Excel Macro files", "*.xlsm")])
        if file:
            entry_cust_horis.delete(0, tk.END)
            entry_cust_horis.insert(0, file)

    def select_hmo_cust():
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        if file:
            entry_hmo_cust.delete(0, tk.END)
            entry_hmo_cust.insert(0, file)

    def select_output_folder():
        folder = filedialog.askdirectory()
        if folder:
            entry_output_folder.delete(0, tk.END)
            entry_output_folder.insert(0, folder)

    def confirm():
        nonlocal root
        # Get values
        in_excel = entry_input_excel.get().strip()
        cust_file = entry_cust_horis.get().strip()
        hmo_file = entry_hmo_cust.get().strip()
        out_folder = entry_output_folder.get().strip()
        sel_date = date_entry.get_date()

        # Validation
        if not in_excel:
            messagebox.showerror("Error", "Please select Input Excel File")
            return
        if not cust_file:
            messagebox.showerror("Error", "Please select Cust Horis .xlsm File")
            return
        if not hmo_file:
            messagebox.showerror("Error", "Please select HMO Cust File")
            return
        if not out_folder:
            messagebox.showerror("Error", "Please select Output Folder")
            return
        if not sel_date:
            messagebox.showerror("Error", "Please select Date")
            return

        # Assign to globals
        globals()["input_excel"] = in_excel
        globals()["cust_horis"] = cust_file
        globals()["hmo_cust"] = hmo_file
        globals()["output_folder"] = out_folder
        globals()["selected_date"] = sel_date

        root.destroy()

    # GUI Window
    root = tk.Tk()
    root.title("Select Files & Parameters")
    root.geometry("600x350")

    # Input Excel
    tk.Label(root, text="Input Excel File:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    entry_input_excel = tk.Entry(root, width=50)
    entry_input_excel.grid(row=0, column=1, padx=10)
    tk.Button(root, text="Browse", command=select_input_excel).grid(row=0, column=2, padx=5)

    # Cust Horis
    tk.Label(root, text="Cust Horis .xlsm File:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    entry_cust_horis = tk.Entry(root, width=50)
    entry_cust_horis.grid(row=1, column=1, padx=10)
    tk.Button(root, text="Browse", command=select_cust_horis).grid(row=1, column=2, padx=5)

    # HMO Cust
    tk.Label(root, text="HMO Cust File:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    entry_hmo_cust = tk.Entry(root, width=50)
    entry_hmo_cust.grid(row=2, column=1, padx=10)
    tk.Button(root, text="Browse", command=select_hmo_cust).grid(row=2, column=2, padx=5)

    # Output Folder
    tk.Label(root, text="Output Folder:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
    entry_output_folder = tk.Entry(root, width=50)
    entry_output_folder.grid(row=3, column=1, padx=10)
    tk.Button(root, text="Browse", command=select_output_folder).grid(row=3, column=2, padx=5)

    # Date Picker
    tk.Label(root, text="Select Date:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
    date_entry = DateEntry(root, width=20, background='darkblue',
                           foreground='white', borderwidth=2, year=2025)
    date_entry.grid(row=4, column=1, padx=10, pady=5, sticky="w")

    # Confirm Button
    tk.Button(root, text="Confirm & Close", command=confirm, bg="lightgreen").grid(row=5, column=1, pady=20)

    root.mainloop()


# Run the GUI
run_gui()

# Now these variables are available outside the GUI
print("Input Excel:", input_excel)
print("Cust Horis:", cust_horis)
print("HMO Cust:", hmo_cust)
print("Output Folder:", output_folder)
print("Selected Date:", selected_date)

```
