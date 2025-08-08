```python
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows

class ReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Reconciliation Tool")
        self.df_original = None
        self.df_transformed = None
        self.selected_columns = []
        self.selected_values = []
        self.remaining_columns = []
        self.numeric_columns = []
        self.current_df = None
        self.init_page1()

    def init_page1(self):
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Button(frame, text="Load Original File", command=lambda: self.load_file('original')).pack(pady=5)
        tk.Button(frame, text="Load Transformed File", command=lambda: self.load_file('transformed')).pack(pady=5)
        tk.Button(frame, text="Next", command=self.page2).pack(pady=5)

    def load_file(self, file_type):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return
        df = pd.read_excel(file_path)

        # Ensure numeric detection
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='ignore')

        if file_type == 'original':
            self.df_original = df
        else:
            self.df_transformed = df

        messagebox.showinfo("Success", f"{file_type.capitalize()} file loaded successfully!")

    def page2(self):
        if self.df_original is None or self.df_transformed is None:
            messagebox.showerror("Error", "Please load both files first.")
            return

        # Detect string columns
        string_columns = self.df_original.select_dtypes(include=['object']).columns.tolist()
        self.remaining_columns = string_columns.copy()

        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(frame)
        scrollbar_y = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollbar_x = tk.Scrollbar(frame, orient="horizontal", command=canvas.xview)
        scroll_frame = tk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        tk.Label(scroll_frame, text="Select String Columns for Grouping").pack(pady=5)
        self.col_vars = {}
        for col in string_columns:
            var = tk.BooleanVar()
            cb = tk.Checkbutton(scroll_frame, text=col, variable=var)
            cb.pack(anchor='w')
            self.col_vars[col] = var

        tk.Button(scroll_frame, text="Submit", command=self.generate_table).pack(pady=5)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

    def generate_table(self):
        self.selected_columns = [col for col, var in self.col_vars.items() if var.get()]
        if not self.selected_columns:
            messagebox.showerror("Error", "Please select at least one column.")
            return

        self.remaining_columns = [col for col in self.remaining_columns if col not in self.selected_columns]

        self.df_original['_filter_key'] = self.df_original[self.selected_columns].astype(str).agg(' | '.join, axis=1)
        self.df_transformed['_filter_key'] = self.df_transformed[self.selected_columns].astype(str).agg(' | '.join, axis=1)

        num_cols = self.df_original.select_dtypes(include=[np.number]).columns.tolist()

        agg_dict = {col: 'sum' for col in num_cols if col != '_filter_key'}
        grouped_orig = self.df_original.groupby('_filter_key', as_index=False).agg(agg_dict)
        grouped_trans = self.df_transformed.groupby('_filter_key', as_index=False).agg(agg_dict)

        merged = pd.merge(grouped_orig, grouped_trans, on='_filter_key', suffixes=('_orig', '_trans'), how='outer')

        merged['Status'] = np.where(merged.isna().any(axis=1), 'Missing',
                                    np.where((merged.filter(like='_orig').values == merged.filter(like='_trans').values).all(axis=1),
                                             'OK', 'Anomaly'))

        self.current_df = merged
        self.show_table()

    def show_table(self):
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True)

        filter_label = tk.Label(frame, text="Filter by _filter_key:")
        filter_label.pack()
        filter_box = ttk.Combobox(frame, values=self.current_df['_filter_key'].unique().tolist())
        filter_box.pack()

        table_frame = tk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True)

        tree = ttk.Treeview(table_frame, columns=list(self.current_df.columns), show='headings')
        for col in self.current_df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        for _, row in self.current_df.iterrows():
            tree.insert("", "end", values=list(row))
        tree.pack(fill=tk.BOTH, expand=True)

        tk.Button(frame, text="Next Drill Down", command=self.page_next).pack(pady=5)
        tk.Button(frame, text="Export to Excel", command=self.export_to_excel).pack(pady=5)

    def page_next(self):
        if not self.remaining_columns:
            messagebox.showinfo("Info", "No more columns to select.")
            return
        # Next selection page logic would be similar to page2
        pass

    def export_to_excel(self):
        if self.current_df is None:
            messagebox.showerror("Error", "No data to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        wb = Workbook()
        ws_data = wb.active
        ws_data.title = "Reconciliation"

        for r in dataframe_to_rows(self.current_df, index=False, header=True):
            ws_data.append(r)

        # Charts per string column
        for col in self.selected_columns:
            ws_chart = wb.create_sheet(title=f"Chart_{col}")
            unique_vals = self.df_original[col].unique().tolist()

            for val in unique_vals:
                chart = BarChart()
                chart.title = f"{col} - {val}"
                chart.x_axis.title = "Metric"
                chart.y_axis.title = "Value"

                data_subset = self.current_df[self.current_df['_filter_key'].str.contains(str(val))]
                if not data_subset.empty:
                    ws_chart.append(["Metric", "Original", "Transformed"])
                    for _, r in data_subset.iterrows():
                        for num_col in self.numeric_columns:
                            ws_chart.append([num_col, r[f"{num_col}_orig"], r[f"{num_col}_trans"]])

                    data_ref = Reference(ws_chart, min_col=2, min_row=1, max_col=3,
                                         max_row=ws_chart.max_row)
                    cats_ref = Reference(ws_chart, min_col=1, min_row=2, max_row=ws_chart.max_row)
                    chart.add_data(data_ref, titles_from_data=True)
                    chart.set_categories(cats_ref)
                    ws_chart.add_chart(chart, "E5")

        wb.save(file_path)
        messagebox.showinfo("Success", "Exported to Excel with charts successfully.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReconciliationApp(root)
    root.mainloop()
```
