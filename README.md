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

        # DataFrames full loaded
        self.df_original_full = None
        self.df_transformed_full = None

        # DataFrames filtered per drill-down step
        self.df_original_filtered = None
        self.df_transformed_filtered = None

        # Columns selected by user
        self.selected_string_cols = []
        self.selected_numeric_cols = []

        # Dictionary of {string_col: list of selected filter values}
        self.current_filters = {}

        self.output_folder = None
        self.export_wb_path = None
        self.export_wb = None
        self.export_counter = 0

        # GUI state holders
        self.current_string_vars = []
        self.current_numeric_vars = []
        self.current_value_listbox = None
        self.current_table_df = None

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
            self.df_original_full = pd.read_excel(orig_path)
            self.df_transformed_full = pd.read_excel(trans_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel files:\n{e}")
            return

        self.df_original_full = self.clean_string_columns(self.df_original_full)
        self.df_transformed_full = self.clean_string_columns(self.df_transformed_full)

        def drop_single_unique_cols(df):
            return df.loc[:, df.apply(lambda col: col.nunique(dropna=True) > 1)]

        df_orig_clean = drop_single_unique_cols(self.df_original_full)
        df_trans_clean = drop_single_unique_cols(self.df_transformed_full)

        common_cols = list(set(df_orig_clean.columns).intersection(set(df_trans_clean.columns)))

        df_orig_common = self.df_original_full[common_cols]
        df_trans_common = self.df_transformed_full[common_cols]

        self.string_cols_available = [col for col in common_cols if df_orig_common[col].dtype == 'object']
        self.numeric_cols_available = [col for col in common_cols if pd.api.types.is_numeric_dtype(df_orig_common[col])]

        if not self.string_cols_available:
            messagebox.showerror("Error", "No common string columns found for selection.")
            return
        if not self.numeric_cols_available:
            messagebox.showerror("Error", "No common numeric columns found for selection.")
            return

        self.selected_string_cols = []
        self.selected_numeric_cols = []
        self.current_filters = {}
        self.export_counter = 0
        self.export_wb_path = None
        self.export_wb = None
        self.output_folder = output_folder

        self.string_col_selection_page(initial=True)

    def string_col_selection_page(self, initial=False):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        title_text = "Select String Columns (Keys for Matching) - Select 1 or more:"
        if not initial:
            title_text = "Select Additional String Columns for Drill Down - Select 1 or more:"

        ttk.Label(frame, text=title_text).pack(anchor="w")

        available = [c for c in self.string_cols_available if c not in self.selected_string_cols]
        if not available:
            # No more columns left to select for drill-down
            self.numeric_col_selection_page()
            return

        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.current_string_vars = []
        for col in available:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(scrollable_frame, text=col, variable=var)
            chk.pack(anchor="w", padx=5, pady=2)
            self.current_string_vars.append((col, var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Next", command=lambda: self.after_string_selection(initial)).pack(side="left", padx=5)
        if initial:
            ttk.Button(btn_frame, text="Back to File Selection", command=self.file_select_screen).pack(side="left", padx=5)
        else:
            ttk.Button(btn_frame, text="Back to Table", command=self.show_table_with_value_filter).pack(side="left", padx=5)

    def after_string_selection(self, initial):
        selected = [col for col, var in self.current_string_vars if var.get()]
        if not selected:
            messagebox.showerror("Selection Error", "Select at least one string column.")
            return
        for col in selected:
            if col not in self.selected_string_cols:
                self.selected_string_cols.append(col)

        if initial:
            self.numeric_col_selection_page()
        else:
            # Drill down - show table with filter by selected values in last selected string col
            self.show_table_with_value_filter()

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
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
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
        ttk.Button(btn_frame, text="Back to String Selection", command=lambda: self.string_col_selection_page(initial=True)).pack(side="left")

    def apply_filters(self):
        """Apply current_filters dict to full DataFrames and update filtered DataFrames"""
        df_orig = self.df_original_full
        df_trans = self.df_transformed_full

        if not self.current_filters:
            self.df_original_filtered = df_orig.copy()
            self.df_transformed_filtered = df_trans.copy()
            return

        mask_orig = pd.Series(True, index=df_orig.index)
        mask_trans = pd.Series(True, index=df_trans.index)

        for col, accepted_values in self.current_filters.items():
            mask_orig &= df_orig[col].isin(accepted_values)
            mask_trans &= df_trans[col].isin(accepted_values)

        self.df_original_filtered = df_orig.loc[mask_orig].copy()
        self.df_transformed_filtered = df_trans.loc[mask_trans].copy()

    def generate_table(self):
        # Get selected numeric cols
        selected_numerics = [col for col, var in self.current_numeric_vars if var.get()]
        if not selected_numerics:
            messagebox.showerror("Selection Error", "Select at least one numeric column.")
            return

        self.selected_numeric_cols = selected_numerics

        # Clear any filters for first table generation
        self.current_filters = {}

        self.apply_filters()

        merged = self.aggregate_and_merge()

        self.current_table_df = merged

        self.show_table_with_value_filter(initial=True)

    def aggregate_and_merge(self):
        group_cols = self.selected_string_cols

        orig_grouped = self.df_original_filtered.groupby(group_cols)[self.selected_numeric_cols].sum().reset_index()
        trans_grouped = self.df_transformed_filtered.groupby(group_cols)[self.selected_numeric_cols].sum().reset_index()

        merged = pd.merge(orig_grouped, trans_grouped, on=group_cols, how='outer', suffixes=('_orig', '_trans'))

        # Fill NaNs in numeric cols with 0 for sum comparisons
        for col in self.selected_numeric_cols:
            merged[f"{col}_orig"] = pd.to_numeric(merged[f"{col}_orig"], errors='coerce').fillna(0)
            merged[f"{col}_trans"] = pd.to_numeric(merged[f"{col}_trans"], errors='coerce').fillna(0)

            def anomaly_status(row):
                v1 = row[f"{col}_orig"]
                v2 = row[f"{col}_trans"]
                if v1 != v2:
                    return "Anomaly"
                return "OK"

            merged[f"{col}_status"] = merged.apply(anomaly_status, axis=1)

        return merged

    def show_table_with_value_filter(self, initial=False):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Show the aggregated table
        if self.current_table_df is None or self.current_table_df.empty:
            ttk.Label(frame, text="No data to display for current filter.").pack()
            btn_frame = ttk.Frame(self.root)
            btn_frame.pack(pady=10)
            ttk.Button(btn_frame, text="Back to String Selection", command=lambda: self.string_col_selection_page(initial=True)).pack()
            return

        # Show filter UI only if drill down possible (more string columns left)
        remaining_string_cols = [c for c in self.string_cols_available if c not in self.selected_string_cols]
        last_selected_col = None
        if self.selected_string_cols:
            last_selected_col = self.selected_string_cols[-1]

        # If drill down possible, show listbox to select values in last selected string col to filter next
        if remaining_string_cols and last_selected_col:
            ttk.Label(frame, text=f"Select values in '{last_selected_col}' to filter drill down (multiple selection allowed):").pack(anchor="w")

            values = sorted(self.current_table_df[last_selected_col].dropna().unique())

            self.current_value_listbox = tk.Listbox(frame, selectmode="extended", height=8)
            for val in values:
                self.current_value_listbox.insert(tk.END, val)
            self.current_value_listbox.pack(fill="both", expand=False)

        else:
            self.current_value_listbox = None

        # Show the grouped aggregated table in Treeview
        table_frame = ttk.Frame(self.root)
        table_frame.pack(padx=10, pady=10, fill="both", expand=True)

        tree = ttk.Treeview(table_frame, show="headings")
        tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)

        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        cols = list(self.current_table_df.columns)

        tree["columns"] = cols

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=130, anchor='w')

        for _, row in self.current_table_df.iterrows():
            values = [row[c] for c in cols]
            tree.insert("", "end", values=values)

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        if self.current_value_listbox is not None:
            ttk.Button(btn_frame, text="Apply Filter & Drill Down", command=self.apply_filter_and_drill_down).pack(side="left", padx=5)

        ttk.Button(btn_frame, text="Export Current Table", command=self.export_current_table).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Start Over (File Selection)", command=self.file_select_screen).pack(side="left", padx=5)

        if initial:
            ttk.Button(btn_frame, text="Back to Numeric Selection", command=self.numeric_col_selection_page).pack(side="left", padx=5)

    def apply_filter_and_drill_down(self):
        # Get selected values in last selected string col to add to filters
        if not self.selected_string_cols:
            messagebox.showerror("Error", "No string columns selected.")
            return

        last_col = self.selected_string_cols[-1]

        if self.current_value_listbox:
            selected_indices = self.current_value_listbox.curselection()
            if not selected_indices:
                messagebox.showerror("Error", "Select at least one value to drill down.")
                return

            selected_values = [self.current_value_listbox.get(i) for i in selected_indices]

            # Update filters dict for last_col
            self.current_filters[last_col] = selected_values

        # Now show next string col selection for drill down
        self.string_col_selection_page(initial=False)

    def export_current_table(self):
        if self.current_table_df is None or self.current_table_df.empty:
            messagebox.showerror("Error", "No table to export.")
            return

        if self.export_wb is None:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            fname = f"Reconcile_{timestamp}.xlsx"
            self.export_wb_path = os.path.join(self.output_folder, fname)
            self.export_wb = Workbook()
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
            messagebox.showinfo("Exported", f"Workbook saved/appended:\n{self.export_wb_path}")
        except PermissionError:
            messagebox.showerror("Error", "Failed to save workbook. Please close it if open and retry.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save workbook:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1200x700")
    app = ReconciliationApp(root)
    root.mainloop()

```
