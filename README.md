```
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

START_ROW = 10  # start reading from this row (10th row)
OUTPUT_SHEET_NAME = "Consolidated"

def consolidate_excel(file_path):
    # Read all sheet names
    xls = pd.ExcelFile(file_path)
    all_dfs = []

    for sheet in xls.sheet_names:
        # Read each sheet starting from row 10 (header = row 10)
        try:
            df = pd.read_excel(file_path, sheet_name=sheet, header=START_ROW - 1)
            df["Source_Sheet"] = sheet  # add sheet name for traceability
            all_dfs.append(df)
        except Exception as e:
            print(f"Skipping {sheet} due to error: {e}")

    # Concatenate all sheets
    if not all_dfs:
        print("No data found to consolidate.")
        return None

    consolidated_df = pd.concat(all_dfs, ignore_index=True)
    return consolidated_df


def main():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo("Select File", "Choose the Excel file to consolidate.")
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
    )

    if not file_path:
        messagebox.showinfo("Cancelled", "No file selected.")
        return

    consolidated_df = consolidate_excel(file_path)
    if consolidated_df is None:
        messagebox.showinfo("No Data", "No valid sheets to consolidate.")
        return

    # Save the consolidated data
    save_path = filedialog.asksaveasfilename(
        title="Save Consolidated File As",
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")]
    )
    if not save_path:
        return

    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
        consolidated_df.to_excel(writer, index=False, sheet_name=OUTPUT_SHEET_NAME)

    messagebox.showinfo("Success", f"Consolidated file saved at:\n{save_path}")


if __name__ == "__main__":
    main()
