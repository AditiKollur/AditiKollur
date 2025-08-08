```python
# Full Python Tkinter GUI for Data Reconciliation with drill-down and chart export
# Features:
#  - Select string columns and numeric columns separately
#  - Multi-select with scrollbars
#  - Filter, group, and aggregate
#  - Drill-down on filtered rows
#  - Export to Excel with one bar chart per numeric column

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
import tempfile
import os

class DataReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Reconciliation Tool")
        self.df_original = None
        self.df_transformed = None
        self.string_cols = []
        self.numeric_cols = []
        self.selected_string_cols = []
        self.selected_numeric_cols = []
        self.selected_values = []
        self.init_file_selection_page()

    def init_file_selection_page(self):
        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text="Select Original File:").pack()
        self.original_file_btn = tk.Button(frame, text="Browse", command=self.load_original_file)
        self.original_file_btn.pack()
        tk.Label(frame, text="Select Transformed File:").pack()
        self.transformed_file_btn = tk.Button(frame, text="Browse", command=self.load_transformed_file)
        self.transformed_file_btn.pack()
        tk.Button(frame, text="Next", command=self.init_column_selection_page).pack(pady=10)

    def load_original_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.df_original = pd.read_excel(file_path)
            messagebox.showinfo("Loaded", "Original file loaded successfully!")

    def load_transformed_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.df_transformed = pd.read_excel(file_path)
            messagebox.showinfo("Loaded", "Transformed file loaded successfully!")

    def init_column_selection_page(self):
        if self.df_original is None or self.df_transformed is None:
            messagebox.showerror("Error", "Please load both files before proceeding.")
            return

        # Get common columns
        common_cols = list(set(self.df_original.columns) & set(self.df_transformed.columns))
        self.string_cols = [c for c in common_cols if self.df_original[c].dtype == 'object']
        self.numeric_cols = [c for c in common_cols if pd.api.types.is_numeric_dtype(self.df_original[c])]

        frame = tk.Frame(self.root)
        for widget in self.root.winfo_children():
            widget.destroy()
        frame.pack(fill="both", expand=True)

        # String columns selection
        tk.Label(frame, text="Select String Columns:").grid(row=0, column=0)
        self.string_listbox = tk.Listbox(frame, selectmode="multiple", exportselection=False, height=10)
        for col in self.string_cols:
            self.string_listbox.insert(tk.END, col)
        self.string_listbox.grid(row=1, column=0, sticky="nsew")
        scrollbar1 = tk.Scrollbar(frame, command=self.string_listbox.yview)
        scrollbar1.grid(row=1, column=1, sticky="ns")
        self.string_listbox.config(yscrollcommand=scrollbar1.set)

        # Numeric columns selection
        tk.Label(frame, text="Select Numeric Columns:").grid(row=0, column=2)
        self.numeric_listbox = tk.Listbox(frame, selectmode="multiple", exportselection=False, height=10)
        for col in self.numeric_cols:
            self.numeric_listbox.insert(tk.END, col)
        self.numeric_listbox.grid(row=1, column=2, sticky="nsew")
        scrollbar2 = tk.Scrollbar(frame, command=self.numeric_listbox.yview)
        scrollbar2.grid(row=1, column=3, sticky="ns")
        self.numeric_listbox.config(yscrollcommand=scrollbar2.set)

        tk.Button(frame, text="Submit", command=self.generate_table).grid(row=2, column=0, columnspan=4, pady=10)

    def generate_table(self):
        self.selected_string_cols = [self.string_cols[i] for i in self.string_listbox.curselection()]
        self.selected_numeric_cols = [self.numeric_cols[i] for i in self.numeric_listbox.curselection()]

        if not self.selected_string_cols or not self.selected_numeric_cols:
            messagebox.showerror("Error", "Please select at least one string and one numeric column.")
            return

        # Merge both DataFrames
        merged = pd.merge(self.df_original, self.df_transformed, on=self.selected_string_cols, suffixes=('_orig', '_trans'))

        # Create filter key
        merged['_filter_key'] = merged[self.selected_string_cols].astype(str).agg(' | '.join, axis=1)

        # Aggregate numeric columns
        agg_dict = {col+'_orig': 'sum' for col in self.selected_numeric_cols}
        agg_dict.update({col+'_trans': 'sum' for col in self.selected_numeric_cols})
        grouped = merged.groupby('_filter_key').agg(agg_dict).reset_index()

        # Detect anomalies
        for col in self.selected_numeric_cols:
            grouped[col+'_status'] = grouped.apply(lambda row: 'OK' if row[col+'_orig'] == row[col+'_trans'] 
                                                   else ('Missing' if row[col+'_trans'] == 0 else 'Anomaly'), axis=1)

        self.show_table(grouped)

    def show_table(self, df):
        for widget in self.root.winfo_children():
            widget.destroy()
        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, columns=list(df.columns), show="headings")
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        for _, row in df.iterrows():
            tree.insert("", tk.END, values=list(row))
        tree.pack(fill="both", expand=True)

        # Scrollbars
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        vsb.pack(side='right', fill='y')
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        hsb.pack(side='bottom', fill='x')
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tk.Button(frame, text="Export to Excel", command=lambda: self.export_to_excel(df)).pack(pady=10)

    def export_to_excel(self, df):
        wb = Workbook()
        ws = wb.active
        ws.title = "Reconciliation Data"
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # One chart per numeric column
        for col in self.selected_numeric_cols:
            chart_ws = wb.create_sheet(title=f"Chart_{col}")
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            df_chart = df[['_filter_key', col+'_orig', col+'_trans']]

            plt.figure(figsize=(8, 5))
            df_chart.set_index('_filter_key')[[col+'_orig', col+'_trans']].plot(kind='bar')
            plt.title(f"Original vs Transformed - {col}")
            plt.tight_layout()
            plt.savefig(temp_file.name)
            plt.close()

            img = Image(temp_file.name)
            chart_ws.add_image(img, "A1")
            os.unlink(temp_file.name)

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            wb.save(file_path)
            messagebox.showinfo("Success", "Excel file saved successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataReconciliationApp(root)
    root.mainloop()
```
