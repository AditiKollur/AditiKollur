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

        self.df_original_full = None
        self.df_transformed_full = None

        self.dfs_original = {}  # dict to hold original dfs per drill down level, keys: d1, d2, ...
        self.dfs_transformed = {}  # same for transformed

        self.all_selected_string_cols = []
        self.selected_numeric_cols = []

        self.output_folder = None
        self.export_wb_path = None
        self.export_wb = None
        self.export_counter = 0

        self.current_filter_level = 0
        self.current_filtered_keys = None  # selected keys from previous drill step

        self.current_string_selection_vars = []
        self.current_numeric_selection_vars = []

        self.current_filter_listbox = None
        self.current_table_df = None
        self.current_filter_key_col = None

        self.string_cols_available = []
        self.numeric_cols_available = []

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

        string_cols = [col for col in common_cols if df_orig_common[col].dtype == 'object']
        numeric_cols = [col for col in common_cols if pd.api.types.is_numeric_dtype(df_orig_common[col])]

        if not string_cols:
            messagebox.showerror("Error", "No common string columns found for selection.")
            return
        if not numeric_cols:
            messagebox.showerror("Error", "No common numeric columns found for selection.")
            return

        self.string_cols_available = string_cols
        self.numeric_cols_available = numeric_cols

        self.all_selected_string_cols = []
        self.selected_numeric_cols = []
        self.export_counter = 0
        self.export_wb_path = None
        self.export_wb = None
        self.output_folder = output_folder

        self.current_filter_level = 0
        self.current_filtered_keys = None
        self.current_table_df = None

        # Initialize dfs at level 1:
        self.dfs_original['d1'] = self.df_original_full.copy()
        self.dfs_transformed['d1'] = self.df_transformed_full.copy()

        self.string_col_selection_page(initial=True)

    def string_col_selection_page(self, initial=False):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        title_text = "Select String Columns (Keys for Matching) - Select 1 or more:"
        if not initial:
            title_text = f"Select Additional String Columns for Drill Down (Level {self.current_filter_level + 1}) - Select 1 or more:"

        ttk.Label(frame, text=title_text).pack(anchor="w")

        available = [c for c in self.string_cols_available if c not in self.all_selected_string_cols]
        if not available:
            messagebox.showinfo("No More Strings", "No more string columns left to select for drill down.\nProceeding to numeric column selection.")
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

        self.current_string_selection_vars = []
        for col in available:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(scrollable_frame, text=col, variable=var)
            chk.pack(anchor="w", padx=5, pady=2)
            self.current_string_selection_vars.append((col, var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Next", command=lambda: self.after_string_selection(initial)).pack(side="left", padx=5)

        if initial:
            ttk.Button(btn_frame, text="Back to File Selection", command=self.file_select_screen).pack(side="left", padx=5)
        else:
            ttk.Button(btn_frame, text="Back to Previous Table", command=self.show_table_with_filters).pack(side="left", padx=5)

    def after_string_selection(self, initial):
        selected_strings = [col for col, var in self.current_string_selection_vars if var.get()]
        if not selected_strings:
            messagebox.showerror("Selection Error", "Select at least one string column.")
            return

        for col in selected_strings:
            if col not in self.all_selected_string_cols:
                self.all_selected_string_cols.append(col)

        if initial:
            self.numeric_col_selection_page()
        else:
            self.generate_table()

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

        self.current_numeric_selection_vars = []
        for col in self.numeric_cols_available:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(scrollable_frame, text=col, variable=var)
            chk.pack(anchor="w", padx=5, pady=2)
            self.current_numeric_selection_vars.append((col, var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        def on_generate():
            selected_numerics = [col for col, var in self.current_numeric_selection_vars if var.get()]
            if not selected_numerics:
                messagebox.showerror("Selection Error", "Select at least one numeric column.")
                return
            self.selected_numeric_cols = selected_numerics
            self.generate_table()

        ttk.Button(btn_frame, text="Generate Table", command=on_generate).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Back to String Selection", command=lambda: self.string_col_selection_page(initial=True)).pack(side="left")

    def generate_table(self):
        if not self.selected_numeric_cols:
            messagebox.showerror("Error", "No numeric columns selected.")
            return

        level_str = f"d{self.current_filter_level + 1}"

        # Filter dfs based on previous selected keys for this level
        if self.current_filter_level == 0:
            df_orig_filtered = self.df_original_full.copy()
            df_trans_filtered = self.df_transformed_full.copy()
        else:
            prev_filter_key_col = f"_filter_key{self.current_filter_level}"
            df_orig_filtered = self.dfs_original[f"d{self.current_filter_level}"]
            df_trans_filtered = self.dfs_transformed[f"d{self.current_filter_level}"]

            if self.current_filtered_keys:
                df_orig_filtered = df_orig_filtered[df_orig_filtered[prev_filter_key_col].isin(self.current_filtered_keys)].copy()
                df_trans_filtered = df_trans_filtered[df_trans_filtered[prev_filter_key_col].isin(self.current_filtered_keys)].copy()

        if df_orig_filtered.empty or df_trans_filtered.empty:
            messagebox.showerror("No Data", "No data available after filtering with current selection.")
            return

        # Create new filter key for current level by concatenating string columns selected so far
        current_filter_key_col = f"_filter_key{self.current_filter_level + 1}"
        cols_to_concat = self.all_selected_string_cols[: self.current_filter_level + 1]

        df_orig_filtered[current_filter_key_col] = df_orig_filtered[cols_to_concat].astype(str).agg(' | '.join, axis=1)
        df_trans_filtered[current_filter_key_col] = df_trans_filtered[cols_to_concat].astype(str).agg(' | '.join, axis=1)

        # Store dfs for this level
        self.dfs_original[level_str] = df_orig_filtered
        self.dfs_transformed[level_str] = df_trans_filtered

        self.current_filter_key_col = current_filter_key_col

        # Group by all selected string cols so far and sum numeric cols
        orig_grouped = df_orig_filtered.groupby(cols_to_concat)[self.selected_numeric_cols].sum().reset_index()
        trans_grouped = df_trans_filtered.groupby(cols_to_concat)[self.selected_numeric_cols].sum().reset_index()

        merged = pd.merge(orig_grouped, trans_grouped, on=cols_to_concat, how='outer', suffixes=('_orig', '_trans'))

        # Round numeric, create anomaly columns
        for col in self.selected_numeric_cols:
            c_orig = f"{col}_orig"
            c_trans = f"{col}_trans"
            merged[c_orig] = pd.to_numeric(merged[c_orig], errors='coerce').round(1)
            merged[c_trans] = pd.to_numeric(merged[c_trans], errors='coerce').round(1)

            def anomaly(row):
                v1, v2 = row[c_orig], row[c_trans]
                if pd.isna(v1) or pd.isna(v2):
                    return "Missing"
                if v1 != v2:
                    return "Anomaly"
                return "OK"

            merged[f"{col}_status"] = merged.apply(anomaly, axis=1)

            merged[c_orig] = merged[c_orig].fillna("Missing")
            merged[c_trans] = merged[c_trans].fillna("Missing")

        display_cols = list(cols_to_concat)
        for col in self.selected_numeric_cols:
            display_cols.extend([f"{col}_orig", f"{col}_trans", f"{col}_status"])

        self.current_table_df = merged[display_cols].copy()

        self.show_table_with_filters()

    def show_table_with_filters(self):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        ttk.Label(frame, text=f"Select filter keys for drilldown (Level {self.current_filter_level + 1}) - multiple selection allowed:").pack(anchor="w")

        filter_keys = sorted(self.current_table_df[self.current_filter_key_col].unique())
        self.current_filtered_keys = None

        list_frame = ttk.Frame(frame)
        list_frame.pack(fill="both", expand=True)

        scrollbar_vert = ttk.Scrollbar(list_frame, orient="vertical")
        scrollbar_horz = ttk.Scrollbar(list_frame, orient="horizontal")

        self.current_filter_listbox = tk.Listbox(list_frame, selectmode='extended',
                                                 yscrollcommand=scrollbar_vert.set,
                                                 xscrollcommand=scrollbar_horz.set, height=8)
        for key in filter_keys:
            self.current_filter_listbox.insert(tk.END, key)

        scrollbar_vert.config(command=self.current_filter_listbox.yview)
        scrollbar_horz.config(command=self.current_filter_listbox.xview)

        self.current_filter_listbox.pack(side="left", fill="both", expand=True)
        scrollbar_vert.pack(side="right", fill="y")
        scrollbar_horz.pack(side="bottom", fill="x")

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
            tree.column(c, width=120, anchor='w')

        for _, row in self.current_table_df.iterrows():
            tree.insert("", "end", values=list(row))

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Drill Down Next", command=self.drill_down_next).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Export Current Table", command=self.export_current_table).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Start Over (File Selection)", command=self.file_select_screen).pack(side="left", padx=5)

    def drill_down_next(self):
        if self.current_filter_listbox:
            selected = [self.current_filter_listbox.get(i) for i in self.current_filter_listbox.curselection()]
            if selected:
                self.current_filtered_keys = selected
            else:
                self.current_filtered_keys = None
        else:
            self.current_filtered_keys = None

        self.current_filter_level += 1
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
            # Remove default sheet
            if "Sheet" in self.export_wb.sheetnames:
                std = self.export_wb["Sheet"]
                self.export_wb.remove(std)
            self.export_counter = 1
        else:
            self.export_counter += 1

        sheet_name = f"Sheet{self.export_counter}"
        chart_sheet_name = f"Chart{self.export_counter}"

        ws = self.export_wb.create_sheet(title=sheet_name)

        for r in dataframe_to_rows(self.current_table_df, index=False, header=True):
            ws.append(r)

        # Create bar chart for first numeric col only
        numeric_col = self.selected_numeric_cols[0]

        try:
            cols = list(self.current_table_df.columns)
            idx_cat = cols.index(self.all_selected_string_cols[self.current_filter_level]) + 1 if self.current_filter_level < len(self.all_selected_string_cols) else 1
            idx_orig = cols.index(f"{numeric_col}_orig") + 1
            idx_trans = cols.index(f"{numeric_col}_trans") + 1
        except Exception:
            idx_cat = 1
            idx_orig = 2
            idx_trans = 3

        max_row = ws.max_row

        chart = BarChart()
        chart.title = f"{numeric_col} Orig vs Trans"
        chart.style = 10
        chart.y_axis.title = numeric_col
        chart.x_axis.title = self.all_selected_string_cols[self.current_filter_level] if self.current_filter_level < len(self.all_selected_string_cols) else "Category"

        data_orig = Reference(ws, min_col=idx_orig, min_row=1, max_row=max_row)
        data_trans = Reference(ws, min_col=idx_trans, min_row=1, max_row=max_row)
        cats = Reference(ws, min_col=idx_cat, min_row=2, max_row=max_row)

        chart.add_data(data_orig, titles_from_data=True)
        chart.add_data(data_trans, titles_from_data=True)
        chart.set_categories(cats)
        chart.shape = 4
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True

        chart_ws = self.export_wb.create_sheet(title=chart_sheet_name)
        chart_ws.add_chart(chart, "A1")

        try:
            self.export_wb.save(self.export_wb_path)
            messagebox.showinfo("Export Success", f"Exported to:\n{self.export_wb_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to save workbook:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1100x700")
    app = ReconciliationApp(root)
    root.mainloop()

```
