```
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import PatternFill

# Define distinct tab colors for each file
TAB_COLORS = ["FFC0CB", "90EE90", "87CEEB", "FFFF99", "C0C0C0", "FFD580"]  # pink, light green, sky blue, yellow, grey, light orange

# Sheets to skip or include once
SKIP_SHEETS = ["Sample"]
SINGLE_INSTANCE_SHEETS = ["ReadMe", "Taxonomy Dropdowns"]

def select_files():
    """Open file dialog to select multiple Excel files."""
    files = filedialog.askopenfilenames(
        title="Select Excel Files",
        filetypes=[("Excel Files", "*.xlsx *.xlsm *.xltx *.xltm")]
    )
    return list(files)

def combine_excels_with_colors(files):
    """Combine all sheets from selected Excel files into one, preserving formatting."""
    if not files:
        messagebox.showerror("Error", "No files selected.")
        return

    combined_wb = Workbook()
    combined_ws = combined_wb.active
    combined_ws.title = "Temp"  # Placeholder sheet to remove later

    used_single_instance = set()
    color_index = 0

    for file in files:
        try:
            wb = load_workbook(file)
        except InvalidFileException:
            messagebox.showwarning("Warning", f"Skipping invalid Excel file: {file}")
            continue

        # Assign color (cycle if more files)
        color = TAB_COLORS[color_index % len(TAB_COLORS)]
        color_index += 1

        for sheet_name in wb.sheetnames:
            # Skip "Sample" sheets
            if sheet_name in SKIP_SHEETS:
                continue

            # Skip single-instance sheets already added
            if sheet_name in SINGLE_INSTANCE_SHEETS and sheet_name in used_single_instance:
                continue

            sheet = wb[sheet_name]

            # Mark tab color
            sheet.sheet_properties.tabColor = color

            # Copy sheet to combined workbook
            combined_wb._add_sheet(sheet.copy_worksheet())

            # Track if single-instance sheet added
            if sheet_name in SINGLE_INSTANCE_SHEETS:
                used_single_instance.add(sheet_name)

    # Remove placeholder sheet
    if "Temp" in combined_wb.sheetnames:
        del combined_wb["Temp"]

    # Save combined file
    output_path = filedialog.asksaveasfilename(
        title="Save Combined Excel As",
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")]
    )

    if output_path:
        combined_wb.save(output_path)
        messagebox.showinfo("Success", f"Combined Excel saved successfully:\n{output_path}")
    else:
        messagebox.showinfo("Cancelled", "Operation cancelled.")

def main():
    root = tk.Tk()
    root.withdraw()  # Hide main window

    messagebox.showinfo("Excel Combiner", "Select the Excel files you want to combine.")
    files = select_files()

    if files:
        combine_excels_with_colors(files)

if __name__ == "__main__":
    main()
