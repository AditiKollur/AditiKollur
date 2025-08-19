
```
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


class FileSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Select ICE Data & Mapping File")
        self.root.geometry("500x250")

        # ICE Data file
        tk.Label(root, text="ICE Data File:", font=("Arial", 12)).pack(pady=5)
        self.ice_entry = tk.Entry(root, width=50)
        self.ice_entry.pack(pady=5)
        tk.Button(root, text="Browse", command=self.browse_ice).pack(pady=5)

        # Mapping file
        tk.Label(root, text="Mapping File:", font=("Arial", 12)).pack(pady=5)
        self.map_entry = tk.Entry(root, width=50)
        self.map_entry.pack(pady=5)
        tk.Button(root, text="Browse", command=self.browse_mapping).pack(pady=5)

        # Submit button
        tk.Button(root, text="Submit", command=self.submit, bg="lightblue").pack(pady=20)

        self.ice_file = None
        self.map_file = None
        self.sheetnames = []

    def browse_ice(self):
        file_path = filedialog.askopenfilename(
            title="Select ICE Data File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.ice_entry.delete(0, tk.END)
            self.ice_entry.insert(0, file_path)

    def browse_mapping(self):
        file_path = filedialog.askopenfilename(
            title="Select Mapping File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.map_entry.delete(0, tk.END)
            self.map_entry.insert(0, file_path)

    def submit(self):
        self.ice_file = self.ice_entry.get().strip()
        self.map_file = self.map_entry.get().strip()

        if not self.ice_file or not self.map_file:
            messagebox.showerror("Error", "Please select both files before submitting.")
            return

        # If ICE data file is Excel, open new GUI to select sheets
        if self.ice_file.endswith((".xlsx", ".xls")):
            try:
                xls = pd.ExcelFile(self.ice_file)
                sheet_list = xls.sheet_names
                self.open_sheet_selector(sheet_list)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        else:
            messagebox.showinfo("Info", "ICE Data file is not Excel. No sheets to select.")

    def open_sheet_selector(self, sheet_list):
        # New window for sheet selection
        sheet_win = tk.Toplevel(self.root)
        sheet_win.title("Select Sheets to Consolidate")
        sheet_win.geometry("400x300")

        tk.Label(sheet_win, text="Select sheets:", font=("Arial", 12)).pack(pady=5)

        # Listbox with multiple selection
        listbox = tk.Listbox(sheet_win, selectmode=tk.MULTIPLE, width=40, height=10)
        for sheet in sheet_list:
            listbox.insert(tk.END, sheet)
        listbox.pack(pady=10)

        def save_selection():
            selected = [listbox.get(i) for i in listbox.curselection()]
            if not selected:
                messagebox.showerror("Error", "Please select at least one sheet.")
                return
            self.sheetnames = selected
            messagebox.showinfo("Selected Sheets", f"Sheets selected:\n{', '.join(self.sheetnames)}")
            sheet_win.destroy()

        tk.Button(sheet_win, text="Confirm Selection", command=save_selection, bg="lightgreen").pack(pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    app = FileSelectorApp(root)
    root.mainloop()



import re
import pandas as pd

def convert_cols(df):
    new_cols = []
    for col in df.columns:
        if re.fullmatch(r"\d{6}", str(col)):  # Matches YYYYMM
            # Convert to datetime
            dt = pd.to_datetime(col, format="%Y%m")
            new_col = dt.strftime("%b%y") + "_YTD"  # MMMYY_YTD
            new_cols.append(new_col)
        else:
            new_cols.append(col)  # keep unchanged
    df.columns = new_cols
    return df



import pandas as pd

def calculate_mtd_lag(df):
    # Get all YTD columns
    ytd_cols = [col for col in df.columns if col.endswith("_YTD")]

    # Sort YTD columns chronologically
    ytd_sorted = sorted(ytd_cols, key=lambda x: pd.to_datetime(x.replace("_YTD", ""), format="%b%y"))

    mtd_data = {}

    # Loop starting from the 2nd index (since we need t-1 and t-2)
    for i in range(2, len(ytd_sorted)):
        curr_col = ytd_sorted[i]      # current YTD
        prev_col = ytd_sorted[i-1]    # previous YTD
        prev2_col = ytd_sorted[i-2]   # previous-2 YTD

        # MTD for current month = YTD(t-1) - YTD(t-2)
        mtd_col = curr_col.replace("_YTD", "_MTD")
        mtd_data[mtd_col] = df[prev_col] - df[prev2_col]

    # Add MTD columns to df
    for col, values in mtd_data.items():
        df[col] = values

    return df
```
