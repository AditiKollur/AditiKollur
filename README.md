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

        self.df_original = None
        self.df_transformed = None

        self.all_selected_string_cols = []
        self.selected_numeric_cols = []

        self.output_folder = None
        self.export_wb_path = None
        self.export_wb = None
        self.export_counter = 0

        self.current_string_selection_vars = []
        self.current_numeric_selection_vars = []

        self.current_filter_listbox = None

        self.current_table_df = None
        self.current_filtered_keys = None

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
        self.current_filtered_keys = None
        self.current_table_df = None

        self.df_original = self.df_original_full.copy()
        self.df_transformed = self.df_transformed_full.copy()

        self.string_col_selection_page(initial=True)

    def make_filter_key(self, df, cols):
        existing_cols = [c for c in cols if c in df.columns]
        if not existing_cols:
            return pd.Series([''] * len(df), index=df.index)
        sliced = df[existing_cols].astype(str)
        if isinstance(sliced, pd.Series):
            return sliced
        return sliced.agg(' | '.join, axis=1)

    def string_col_selection_page(self, initial=False):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        title_text = "Select String Columns (Keys for Matching) - Select 1 or more:"
        if not initial:
            title_text = "Select Additional String Columns for Drill Down - Select 1 or more:"

        ttk.Label(frame, text=title_text).pack(anchor="w")

        available = [c for c in self.string_cols_available if c not in self.all_selected_string_cols]
        if not available:
            messagebox.showinfo("No More Strings", "No more string columns left to select for drill down.\nProceeding to export and inspection.")
            self.show_table_with_filters()
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
        ttk.Button(btn_frame, text="Generate Table", command=self.generate_table).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Back to String Selection", command=lambda: self.string_col_selection_page(initial=True)).pack(side="left")

    def generate_table(self):
        if not self.selected_numeric_cols:
            selected_numerics = [col for col, var in self.current_numeric_selection_vars if var.get()]
            if not selected_numerics:
                messagebox.showerror("Selection Error", "Select at least one numeric column.")
                return
            self.selected_numeric_cols = selected_numerics

        # When filtering dfs, use full dfs and filter on _filter_key generated from all_selected_string_cols
        if self.current_filtered_keys:
            mask_orig = self.make_filter_key(self.df_original_full, self.all_selected_string_cols).isin(self.current_filtered_keys)
            mask_trans = self.make_filter_key(self.df_transformed_full, self.all_selected_string_cols).isin(self.current_filtered_keys)
            self.df_original = self.df_original_full.loc[mask_orig].copy()
            self.df_transformed = self.df_transformed_full.loc[mask_trans].copy()
        else:
            self.df_original = self.df_original_full.copy()
            self.df_transformed = self.df_transformed_full.copy()

        # Assign _filter_key safely as Series
        self.df_original['_filter_key'] = self.make_filter_key(self.df_original, self.all_selected_string_cols)
        self.df_transformed['_filter_key'] = self.make_filter_key(self.df_transformed, self.all_selected_string_cols)

        group_cols = self.all_selected_string_cols
        orig_grouped = self.df_original.groupby(group_cols)[self.selected_numeric_cols].sum().reset_index()
        trans_grouped = self.df_transformed.groupby(group_cols)[self.selected_numeric_cols].sum().reset_index()

        merged = pd.merge(orig_grouped, trans_grouped, on=group_cols, how='outer', suffixes=('_orig', '_trans'))

        merged['_filter_key'] = self.make_filter_key(merged, self.all_selected_string_cols)

        for col in self.selected_numeric_cols:
            col_orig = f"{col}_orig"
            col_trans = f"{col}_trans"
            merged[col_orig] = pd.to_numeric(merged[col_orig], errors='coerce').round(1)
            merged[col_trans] = pd.to_numeric(merged[col_trans], errors='coerce').round(1)

            def anomaly_status(row):
                v1, v2 = row[col_orig], row[col_trans]
                if pd.isna(v1) or pd.isna(v2):
                    return "Missing"
                if v1 != v2:
                    return "Anomaly"
                return "OK"

            merged[f"{col}_status"] = merged.apply(anomaly_status, axis=1)
            merged[col_orig] = merged[col_orig].fillna("Missing")
            merged[col_trans] = merged[col_trans].fillna("Missing")

        self.current_table_df = merged.copy()

        self.show_table_with_filters()

    def show_table_with_filters(self):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        ttk.Label(frame, text="Select filter keys for drilldown (multiple selection allowed):").pack(anchor="w")

        filter_keys = sorted(self.current_table_df['_filter_key'].unique())
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

        # Table frame
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
        # Hide _filter_key column from display
        if '_filter_key' in cols:
            cols.remove('_filter_key')
        tree["columns"] = cols

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=120, anchor='w')

        for _, row in self.current_table_df.iterrows():
            values = [row[c] for c in cols]
            tree.insert("", "end", values=values)

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

            chart_ws.add_chart(chart, f"A{1 + 15*i}")

        try:
            self.export_wb.save(self.export_wb_path)
            messagebox.showinfo("Exported", f"Exported table and charts to:\n{self.export_wb_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to save Excel file:\n{e}")



if __name__ == "__main__":
    root = tk.Tk()
    app = ReconciliationApp(root)
    root.geometry("950x650")
    root.mainloop()

```
