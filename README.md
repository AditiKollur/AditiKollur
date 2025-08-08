```python
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import datetime
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList


class ReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Reconciliation")

        # Full original & transformed DataFrames
        self.df_original = None
        self.df_transformed = None

        # User selected columns
        self.selected_string_cols = []
        self.selected_numeric_cols = []

        # Current filtered keys on concatenated key col
        self.current_filter_keys = None  # None means no filter (all rows)

        # Current aggregated DataFrame with concatenated filter key col
        self.current_table_df = None

        self.output_folder = None
        self.export_wb = None
        self.export_wb_path = None
        self.export_counter = 0

        self.file_select_screen()

    def clear_gui(self):
        for w in self.root.winfo_children():
            w.destroy()

    def file_select_screen(self):
        self.clear_gui()
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Select Original Data File (Excel):").grid(row=0, column=0, sticky="w")
        self.orig_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.orig_path_var, width=60).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(frame, text="Browse", command=self.browse_orig_file).grid(row=0, column=2, padx=5)

        ttk.Label(frame, text="Select Transformed Data File (Excel):").grid(row=1, column=0, sticky="w", pady=10)
        self.trans_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.trans_path_var, width=60).grid(row=1, column=1, sticky="ew", padx=5, pady=10)
        ttk.Button(frame, text="Browse", command=self.browse_trans_file).grid(row=1, column=2, padx=5, pady=10)

        ttk.Label(frame, text="Select Output Folder:").grid(row=2, column=0, sticky="w")
        self.output_folder_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.output_folder_var, width=60).grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(frame, text="Browse", command=self.browse_output_folder).grid(row=2, column=2, padx=5)

        frame.columnconfigure(1, weight=1)

        ttk.Button(frame, text="Load Data", command=self.load_data).grid(row=3, column=0, columnspan=3, pady=20)

    def browse_orig_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.orig_path_var.set(path)

    def browse_trans_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.trans_path_var.set(path)

    def browse_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder_var.set(folder)

    def clean_string_columns(self, df):
        str_cols = df.select_dtypes(include=['object']).columns
        for col in str_cols:
            df[col] = df[col].fillna('')
            df[col] = df[col].apply(lambda x: '' if isinstance(x, str) and x.strip() == '' else x)
        return df

    def load_data(self):
        orig_path = self.orig_path_var.get()
        trans_path = self.trans_path_var.get()
        output_folder = self.output_folder_var.get()

        if not orig_path or not os.path.exists(orig_path):
            messagebox.showerror("Error", "Original data file not found or invalid.")
            return
        if not trans_path or not os.path.exists(trans_path):
            messagebox.showerror("Error", "Transformed data file not found or invalid.")
            return
        if not output_folder or not os.path.isdir(output_folder):
            messagebox.showerror("Error", "Output folder invalid or not selected.")
            return

        try:
            self.df_original = pd.read_excel(orig_path)
            self.df_transformed = pd.read_excel(trans_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel files:\n{e}")
            return

        self.df_original = self.clean_string_columns(self.df_original)
        self.df_transformed = self.clean_string_columns(self.df_transformed)

        # Find common columns with variance
        common_cols = list(set(self.df_original.columns).intersection(set(self.df_transformed.columns)))

        # Filter common columns where both have variance
        def cols_with_variance(df, cols):
            return [c for c in cols if df[c].nunique(dropna=True) > 1]

        common_cols_var = list(set(cols_with_variance(self.df_original, common_cols)).intersection(
            set(cols_with_variance(self.df_transformed, common_cols))
        ))

        self.string_cols_available = [c for c in common_cols_var if self.df_original[c].dtype == 'object']
        self.numeric_cols_available = [c for c in common_cols_var if pd.api.types.is_numeric_dtype(self.df_original[c])]

        if not self.string_cols_available:
            messagebox.showerror("Error", "No common string columns found for selection.")
            return
        if not self.numeric_cols_available:
            messagebox.showerror("Error", "No common numeric columns found for selection.")
            return

        self.selected_string_cols = []
        self.selected_numeric_cols = []
        self.current_filter_keys = None
        self.current_table_df = None

        self.export_counter = 0
        self.export_wb = None
        self.export_wb_path = None
        self.output_folder = output_folder

        self.string_col_selection_page()

    def string_col_selection_page(self):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        ttk.Label(frame, text="Select String Columns (Keys for Grouping & Filtering) - Select 1 or more:").pack(anchor="w")

        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.current_string_vars = []
        for col in self.string_cols_available:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(scrollable_frame, text=col, variable=var)
            chk.pack(anchor="w", padx=5, pady=2)
            self.current_string_vars.append((col, var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Next", command=self.after_string_selection).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Back to File Selection", command=self.file_select_screen).pack(side="left", padx=5)

    def after_string_selection(self):
        selected = [col for col, var in self.current_string_vars if var.get()]
        if not selected:
            messagebox.showerror("Selection Error", "Select at least one string column.")
            return
        self.selected_string_cols = selected
        self.numeric_col_selection_page()

    def numeric_col_selection_page(self):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        ttk.Label(frame, text="Select Numeric Columns (Measures) - Select 1 or more:").pack(anchor="w")

        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.current_numeric_vars = []
        for col in self.numeric_cols_available:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(scrollable_frame, text=col, variable=var)
            chk.pack(anchor="w", padx=5, pady=2)
            self.current_numeric_vars.append((col, var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Generate Table", command=self.generate_table).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Back to String Selection", command=self.string_col_selection_page).pack(side="left", padx=5)

    def generate_table(self):
        selected_numerics = [col for col, var in self.current_numeric_vars if var.get()]
        if not selected_numerics:
            messagebox.showerror("Selection Error", "Select at least one numeric column.")
            return

        self.selected_numeric_cols = selected_numerics
        # Start with no filter keys selected (all rows)
        self.current_filter_keys = None

        self.regenerate_table()

    def concat_filter_key(self, df, cols):
        # Concatenate string columns values with '||' separator to form filter key
        return df[cols].astype(str).agg('||'.join, axis=1)

    def regenerate_table(self):
        # If current_filter_keys is None => use all rows, else filter current_table_df by these keys

        # First, we need to get the subset of original and transformed data to aggregate

        # Filter original and transformed data based on current filter keys on previously selected string columns

        if self.current_filter_keys is None:
            # No filter applied => use full data
            df_orig_filtered = self.df_original.copy()
            df_trans_filtered = self.df_transformed.copy()
        else:
            # Filter original and transformed by matching on concatenated string key columns in previous aggregation
            # We must reconstruct concatenated key in original data to filter on keys

            # Create concatenated key in original data
            key_cols = self.selected_string_cols
            df_orig_key = self.concat_filter_key(self.df_original, key_cols)
            df_trans_key = self.concat_filter_key(self.df_transformed, key_cols)

            mask_orig = df_orig_key.isin(self.current_filter_keys)
            mask_trans = df_trans_key.isin(self.current_filter_keys)

            df_orig_filtered = self.df_original.loc[mask_orig].copy()
            df_trans_filtered = self.df_transformed.loc[mask_trans].copy()

        # Now aggregate on selected string columns
        group_cols = self.selected_string_cols

        if len(group_cols) == 0:
            messagebox.showerror("Error", "No string columns selected for grouping.")
            return

        try:
            orig_grouped = df_orig_filtered.groupby(group_cols)[self.selected_numeric_cols].sum(min_count=1).reset_index()
            trans_grouped = df_trans_filtered.groupby(group_cols)[self.selected_numeric_cols].sum(min_count=1).reset_index()
        except Exception as e:
            messagebox.showerror("Error during aggregation:\n" + str(e))
            return

        merged = pd.merge(orig_grouped, trans_grouped, on=group_cols, how='outer', suffixes=('_orig', '_trans'))

        # Fill NaNs in numeric cols with 0
        for col in self.selected_numeric_cols:
            merged[f"{col}_orig"] = pd.to_numeric(merged[f"{col}_orig"], errors='coerce').fillna(0)
            merged[f"{col}_trans"] = pd.to_numeric(merged[f"{col}_trans"], errors='coerce').fillna(0)

            # Anomaly status
            def anomaly(row):
                return "Anomaly" if row[f"{col}_orig"] != row[f"{col}_trans"] else "OK"

            merged[f"{col}_status"] = merged.apply(anomaly, axis=1)

        # Add concatenated filter key column for filtering
        merged['_filter_key'] = self.concat_filter_key(merged, group_cols)

        # Sort by _filter_key for UI ease
        merged.sort_values('_filter_key', inplace=True)

        self.current_table_df = merged

        self.show_table_with_filter()

    def show_table_with_filter(self):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(frame, text="Aggregated Table (with concatenated filter key):").pack(anchor="w")

        # Filter UI for _filter_key column
        ttk.Label(frame, text="Select filter key(s) for drill down (multiple selection allowed):").pack(anchor="w")

        filter_keys = list(self.current_table_df['_filter_key'].unique())

        self.filter_key_listbox = tk.Listbox(frame, selectmode="extended", height=12)
        for key in filter_keys:
            self.filter_key_listbox.insert(tk.END, key)
        self.filter_key_listbox.pack(fill="both", expand=False, pady=5)

        # Show table in Treeview excluding the _filter_key column (can be hidden but showing here for debugging)
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(tree_frame, show="headings")
        tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        # Columns except _filter_key (but you can show it if you want)
        display_cols = [c for c in self.current_table_df.columns if c != '_filter_key']

        tree["columns"] = display_cols

        for col in display_cols:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="w")

        for _, row in self.current_table_df.iterrows():
            values = [row[col] for col in display_cols]
            tree.insert("", "end", values=values)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Drill Down with Selected Filters", command=self.drill_down).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Export Current Table", command=self.export_current_table).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Restart", command=self.file_select_screen).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Add More String Columns & Regenerate", command=self.add_more_string_columns).pack(side="left", padx=5)

    def drill_down(self):
        # Get selected filter keys from listbox
        selected_indices = self.filter_key_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "Select at least one filter key to drill down.")
            return
        selected_keys = [self.filter_key_listbox.get(i) for i in selected_indices]

        self.current_filter_keys = selected_keys

        self.regenerate_table()

    def add_more_string_columns(self):
        # Allow user to select additional string columns (not already selected)
        remaining_cols = [c for c in self.string_cols_available if c not in self.selected_string_cols]
        if not remaining_cols:
            messagebox.showinfo("Info", "No more string columns available to add.")
            return

        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        ttk.Label(frame, text="Select Additional String Columns to Add (for next drill down):").pack(anchor="w")

        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.current_string_vars = []
        for col in remaining_cols:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(scrollable_frame, text=col, variable=var)
            chk.pack(anchor="w", padx=5, pady=2)
            self.current_string_vars.append((col, var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Add & Regenerate Table", command=self.after_add_more_string_columns).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Back to Table", command=self.show_table_with_filter).pack(side="left", padx=5)

    def after_add_more_string_columns(self):
        selected = [col for col, var in self.current_string_vars if var.get()]
        if not selected:
            messagebox.showerror("Selection Error", "Select at least one string column to add.")
            return

        self.selected_string_cols.extend(selected)

        self.regenerate_table()

    def export_current_table(self):
        if self.current_table_df is None or self.current_table_df.empty:
            messagebox.showerror("Error", "No data to export.")
            return

        if self.export_wb is None:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            fname = f"Reconcile_{timestamp}.xlsx"
            self.export_wb_path = os.path.join(self.output_folder, fname)
            self.export_wb = Workbook()
            # Remove default sheet created by openpyxl
            default_sheet = self.export_wb.active
            self.export_wb.remove(default_sheet)
            self.export_counter = 1
        else:
            self.export_counter += 1

        sheet_name = f"Sheet{self.export_counter}"
        chart_sheet_name = f"Chart{self.export_counter}"

        ws = self.export_wb.create_sheet(title=sheet_name)

        for r in dataframe_to_rows(self.current_table_df, index=False, header=True):
            ws.append(r)

        chart_ws = self.export_wb.create_sheet(title=chart_sheet_name)

        for i, col in enumerate(self.selected_numeric_cols):
            col_orig = f"{col}_orig"
            col_trans = f"{col}_trans"
            max_row = ws.max_row
            cats = Reference(ws, min_col=1, min_row=2, max_row=max_row)

            try:
                idx_orig = list(self.current_table_df.columns).index(col_orig) + 1
                idx_trans = list(self.current_table_df.columns).index(col_trans) + 1
            except ValueError:
                continue

            values_orig = Reference(ws, min_col=idx_orig, min_row=2, max_row=max_row)
            values_trans = Reference(ws, min_col=idx_trans, min_row=2, max_row=max_row)

            chart = BarChart()
            chart.title = f"Original vs Transformed - {col}"
            chart.y_axis.title = col
            chart.x_axis.title = "Filter Keys"

            chart.add_data(values_orig, titles_from_data=False, title="Original")
            chart.add_data(values_trans, titles_from_data=False, title="Transformed")
            chart.set_categories(cats)
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showVal = True

            chart_ws.add_chart(chart, f"A{1 + i * 15}")

        try:
            self.export_wb.save(self.export_wb_path)
            messagebox.showinfo("Export Successful", f"Workbook saved/appended:\n{self.export_wb_path}")
        except PermissionError:
            messagebox.showerror("Error", "Close the Excel file if it's open and try again.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save workbook:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1200x700")
    app = ReconciliationApp(root)
    root.mainloop()

```
