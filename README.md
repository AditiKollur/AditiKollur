import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import datetime
import os

from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList


class ReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Reconciliation")

        # DataFrames full original and transformed (loaded once)
        self.df_original_full = None
        self.df_transformed_full = None

        # Filtered working dfs for drill down steps
        self.df_original = None
        self.df_transformed = None

        # Selected string and numeric columns so far (for _filter_key)
        self.all_selected_string_cols = []
        self.selected_numeric_cols = []

        # For export
        self.output_folder = None
        self.export_wb_path = None
        self.export_wb = None
        self.export_counter = 0

        # Cache _filter_key for full dataframes to speed filtering
        self._original_filter_key_cache = None
        self._transformed_filter_key_cache = None

        self.current_string_selection_vars = []
        self.current_numeric_selection_vars = []

        self.current_filter_listbox = None

        # Start from file selection screen
        self.file_select_screen()

    def clear_gui(self):
        for w in self.root.winfo_children():
            w.destroy()

    def file_select_screen(self):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        ttk.Label(frame, text="Select Original Data File (Excel):").pack(anchor="w")
        self.orig_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.orig_path_var, width=50).pack(side="left", fill="x", expand=True)
        ttk.Button(frame, text="Browse", command=self.browse_orig_file).pack(side="left", padx=5)

        ttk.Label(frame, text="Select Transformed Data File (Excel):").pack(anchor="w", pady=(10,0))
        self.trans_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.trans_path_var, width=50).pack(side="left", fill="x", expand=True)
        ttk.Button(frame, text="Browse", command=self.browse_trans_file).pack(side="left", padx=5)

        ttk.Label(frame, text="Select Output Folder:").pack(anchor="w", pady=(10,0))
        self.output_folder_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.output_folder_var, width=50).pack(side="left", fill="x", expand=True)
        ttk.Button(frame, text="Browse", command=self.browse_output_folder).pack(side="left", padx=5)

        ttk.Button(frame, text="Load Data", command=self.load_data).pack(pady=20)

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

        # Drop columns with only 1 unique value or all NA from both dfs before finding common columns
        def drop_single_unique_cols(df):
            return df.loc[:, df.apply(lambda col: col.nunique(dropna=True) > 1)]

        df_orig_clean = drop_single_unique_cols(self.df_original_full)
        df_trans_clean = drop_single_unique_cols(self.df_transformed_full)

        # Find common columns
        common_cols = list(set(df_orig_clean.columns).intersection(set(df_trans_clean.columns)))

        # From common columns, find string and numeric columns
        df_orig_common = self.df_original_full[common_cols]
        df_trans_common = self.df_transformed_full[common_cols]

        # Identify string columns as object dtype, numeric columns by pandas API
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

        # Reset selections and caches
        self.all_selected_string_cols = []
        self.selected_numeric_cols = []
        self.export_counter = 0
        self.export_wb_path = None
        self.export_wb = None
        self.output_folder = output_folder
        self._original_filter_key_cache = None
        self._transformed_filter_key_cache = None

        # Set working dfs initially full
        self.df_original = self.df_original_full.copy()
        self.df_transformed = self.df_transformed_full.copy()

        self.string_col_selection_page()

    def string_col_selection_page(self):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        ttk.Label(frame, text="Select String Columns (Keys for Matching) - Select 1 or more:").pack(anchor="w")

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
        for col in self.string_cols_available:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(scrollable_frame, text=col, variable=var)
            chk.pack(anchor="w", padx=5, pady=2)
            self.current_string_selection_vars.append((col, var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Button(self.root, text="Next", command=self.numeric_col_selection_page).pack(pady=10)
        ttk.Button(self.root, text="Back to File Selection", command=self.file_select_screen).pack()

    def numeric_col_selection_page(self):
        # Collect selected string columns
        selected_strings = [col for col, var in self.current_string_selection_vars if var.get()]
        if not selected_strings:
            messagebox.showerror("Selection Error", "Select at least one string column.")
            return
        self.all_selected_string_cols.extend(selected_strings)

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
        ttk.Button(btn_frame, text="Back to String Selection", command=self.string_col_selection_page).pack(side="left")

    def generate_table(self):
        selected_numerics = [col for col, var in self.current_numeric_selection_vars if var.get()]
        if not selected_numerics:
            messagebox.showerror("Selection Error", "Select at least one numeric column.")
            return

        self.selected_numeric_cols = selected_numerics

        # Prepare working dataframes by filtering original and transformed by accumulated _filter_key if drill down step > 1
        if self._original_filter_key_cache is None:
            self._original_filter_key_cache = self.df_original_full[self.all_selected_string_cols].astype(str).agg(' | '.join, axis=1)
            self._transformed_filter_key_cache = self.df_transformed_full[self.all_selected_string_cols].astype(str).agg(' | '.join, axis=1)

        # Filter full dfs by previous filter if drilldown (if _filter_key exists in current df)
        if hasattr(self, 'current_filtered_keys') and self.current_filtered_keys:
            mask_orig = self._original_filter_key_cache.isin(self.current_filtered_keys)
            mask_trans = self._transformed_filter_key_cache.isin(self.current_filtered_keys)
            self.df_original = self.df_original_full[mask_orig].copy()
            self.df_transformed = self.df_transformed_full[mask_trans].copy()
        else:
            self.df_original = self.df_original_full.copy()
            self.df_transformed = self.df_transformed_full.copy()

        # Generate _filter_key for this step based on currently selected string columns only (concat with previous keys)
        # To maintain filter keys correctly, we always concatenate all selected string columns
        self.df_original['_filter_key'] = self.df_original[self.all_selected_string_cols].astype(str).agg(' | '.join, axis=1)
        self.df_transformed['_filter_key'] = self.df_transformed[self.all_selected_string_cols].astype(str).agg(' | '.join, axis=1)

        # Group and sum numeric columns by _filter_key
        group_cols = ['_filter_key']
        orig_grouped = self.df_original.groupby(group_cols)[selected_numerics].sum().reset_index()
        trans_grouped = self.df_transformed.groupby(group_cols)[selected_numerics].sum().reset_index()

        # Merge on _filter_key
        merged = pd.merge(orig_grouped, trans_grouped, on='_filter_key', how='outer', suffixes=('_orig', '_trans'))

        # Round numeric values to 1 decimal and check anomaly status
        for col in selected_numerics:
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

        # Save this table for filtering in next drill down
        self.current_table_df = merged.copy()

        self.show_table_with_filters()

    def show_table_with_filters(self):
        self.clear_gui()
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Instruction label
        ttk.Label(frame, text="Select filter keys for drilldown (multiple selection allowed):").pack(anchor="w")

        # Filter keys listbox with multiple select
        filter_keys = sorted(self.current_table_df['_filter_key'].unique())
        self.current_filtered_keys = None  # reset current filter keys

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

        self.current_filter_listbox.grid(row=0, column=0, sticky='nsew')
        scrollbar_vert.grid(row=0, column=1, sticky='ns')
        scrollbar_horz.grid(row=1, column=0, sticky='ew')

        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        # Table below showing full merged dataframe

        table_frame = ttk.Frame(frame)
        table_frame.pack(fill="both", expand=True, pady=10)

        columns = list(self.current_table_df.columns)
        tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        tree.pack(side="left", fill="both", expand=True)

        # Scrollbars for treeview
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")

        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="w", width=110)

        for _, row in self.current_table_df.iterrows():
            vals = list(row)
            tree.insert("", "end", values=vals)

        # Buttons
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Drill Down Next Step", command=self.drill_down).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Start Over (File Select)", command=self.file_select_screen).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Export Current Table", command=self.export_current_table).pack(side="left", padx=5)

    def drill_down(self):
        if not self.current_filter_listbox:
            messagebox.showerror("Error", "No filter keys listbox available.")
            return
        selected_indices = self.current_filter_listbox.curselection()
        if selected_indices:
            selected_keys = [self.current_filter_listbox.get(i) for i in selected_indices]
            self.current_filtered_keys = selected_keys
        else:
            # If none selected, use all keys
            self.current_filtered_keys = sorted(self.current_table_df['_filter_key'].unique())

        # Now filter full original and transformed dfs by _filter_key matching selected keys
        # Rebuild caches if not done
        if self._original_filter_key_cache is None:
            self._original_filter_key_cache = self.df_original_full[self.all_selected_string_cols].astype(str).agg(' | '.join, axis=1)
            self._transformed_filter_key_cache = self.df_transformed_full[self.all_selected_string_cols].astype(str).agg(' | '.join, axis=1)

        mask_orig = self._original_filter_key_cache.isin(self.current_filtered_keys)
        mask_trans = self._transformed_filter_key_cache.isin(self.current_filtered_keys)

        self.df_original = self.df_original_full[mask_orig].copy()
        self.df_transformed = self.df_transformed_full[mask_trans].copy()

        # Go back to numeric selection page for next drill step
        self.numeric_col_selection_page()

    def export_current_table(self):
        if not self.output_folder:
            messagebox.showerror("Error", "Output folder not set.")
            return

        if not hasattr(self, 'current_table_df'):
            messagebox.showerror("Error", "No current table to export.")
            return

        if self.export_wb_path is None:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Reconcile_{timestamp}.xlsx"
            self.export_wb_path = os.path.join(self.output_folder, filename)
            self.export_wb = Workbook()
            default_sheet = self.export_wb.active
            self.export_wb.remove(default_sheet)
            self.export_counter = 0
        else:
            try:
                self.export_wb = load_workbook(self.export_wb_path)
            except PermissionError:
                messagebox.showerror("Error", "Workbook is open in another program. Please close and retry.")
                return
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load workbook:\n{e}")
                return

        self.export_counter += 1
        sheet_name = f"Sheet{self.export_counter}"
        chart_sheet_name = f"Chart{self.export_counter}"

        # Add data sheet
        ws = self.export_wb.create_sheet(title=sheet_name)
        for r in dataframe_to_rows(self.current_table_df, index=False, header=True):
            ws.append(r)

        # Add chart sheet
        chart_ws = self.export_wb.create_sheet(title=chart_sheet_name)

        # Create bar chart per numeric col
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
            chart.x_axis.title = "Keys"

            chart.add_data(values_orig, titles_from_data=False, title="Original")
            chart.add_data(values_trans, titles_from_data=False, title="Transformed")
            chart.set_categories(cats)
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showVal = True

            # Position charts vertically spaced
            chart_ws.add_chart(chart, f"A{1 + i*15}")

        try:
            self.export_wb.save(self.export_wb_path)
            messagebox.showinfo("Exported", f"Workbook saved/appended:\n{self.export_wb_path}")
        except PermissionError:
            messagebox.showerror("Error", "Failed to save workbook. Close it if open and retry.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save workbook:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("900x700")
    app = ReconciliationApp(root)
    root.mainloop()
