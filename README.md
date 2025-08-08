```python
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
import tempfile
import os
import datetime


class DataReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Reconciliation Tool")

        self.df_original_full = None
        self.df_transformed_full = None
        self.df_original = None
        self.df_transformed = None

        self.string_cols = []
        self.numeric_cols = []

        self.all_selected_string_cols = []  # cumulatively selected string columns (dimensions)
        self.selected_numeric_cols = []

        self.export_wb_path = None
        self.export_wb = None
        self.export_counter = 0

        self.output_folder = None

        self.init_file_selection_page()

    def clear_root(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def drop_single_unique_columns(self, df):
        # Drop columns with only 1 unique value including NaNs
        return df.loc[:, df.nunique(dropna=False) > 1]

    def init_file_selection_page(self):
        self.clear_root()
        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        tk.Label(frame, text="Select Original File:").pack(anchor='w')
        tk.Button(frame, text="Browse Original File", command=self.load_original_file).pack(pady=(0, 10), anchor='w')

        tk.Label(frame, text="Select Transformed File:").pack(anchor='w')
        tk.Button(frame, text="Browse Transformed File", command=self.load_transformed_file).pack(pady=(0, 10), anchor='w')

        tk.Label(frame, text="Select Output Folder (for saving reconciliation workbook):").pack(anchor='w')
        folder_frame = tk.Frame(frame)
        folder_frame.pack(fill='x', pady=(0, 15))
        self.folder_path_var = tk.StringVar()
        folder_entry = tk.Entry(folder_frame, textvariable=self.folder_path_var, state='readonly')
        folder_entry.pack(side='left', fill='x', expand=True)
        tk.Button(folder_frame, text="Browse Folder", command=self.select_output_folder).pack(side='left', padx=5)

        tk.Button(frame, text="Next", command=self.init_column_selection_page).pack()

        # Reset internal state on restart
        self.df_original_full = None
        self.df_transformed_full = None
        self.df_original = None
        self.df_transformed = None
        self.string_cols = []
        self.numeric_cols = []
        self.all_selected_string_cols = []
        self.selected_numeric_cols = []
        self.export_wb_path = None
        self.export_wb = None
        self.export_counter = 0
        self.output_folder = None
        self.folder_path_var.set('')

    def select_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder = folder_selected
            self.folder_path_var.set(folder_selected)

    def load_original_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.df_original_full = pd.read_excel(file_path)
                messagebox.showinfo("Loaded", "Original file loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{e}")

    def load_transformed_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.df_transformed_full = pd.read_excel(file_path)
                messagebox.showinfo("Loaded", "Transformed file loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{e}")

    def init_column_selection_page(self):
        if self.df_original_full is None or self.df_transformed_full is None:
            messagebox.showerror("Error", "Please load both files before proceeding.")
            return
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder before proceeding.")
            return

        self.df_original = self.df_original_full.copy()
        self.df_transformed = self.df_transformed_full.copy()

        df_orig_trimmed = self.drop_single_unique_columns(self.df_original)
        df_trans_trimmed = self.drop_single_unique_columns(self.df_transformed)

        common_cols = list(set(df_orig_trimmed.columns) & set(df_trans_trimmed.columns))

        self.string_cols = [c for c in common_cols if df_orig_trimmed[c].dtype == 'object']
        self.numeric_cols = [c for c in common_cols if pd.api.types.is_numeric_dtype(df_orig_trimmed[c])]

        # Exclude already selected string columns for drill down, show remaining
        remaining_string_cols = [c for c in self.string_cols if c not in self.all_selected_string_cols]

        self.clear_root()
        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Dimensions (string columns) label + listbox with scrollbars
        tk.Label(frame, text="Select String Columns (Dimensions):").grid(row=0, column=0, sticky='w')

        string_frame = tk.Frame(frame)
        string_frame.grid(row=1, column=0, sticky='nsew', padx=(0,10), pady=5)
        self.string_listbox = tk.Listbox(string_frame, selectmode='multiple', exportselection=False, height=15)
        self.string_listbox.pack(side='left', fill='both', expand=True)

        string_vscroll = ttk.Scrollbar(string_frame, orient='vertical', command=self.string_listbox.yview)
        string_vscroll.pack(side='right', fill='y')
        string_hscroll = ttk.Scrollbar(frame, orient='horizontal', command=self.string_listbox.xview)
        string_hscroll.grid(row=2, column=0, sticky='ew', padx=(0,10))

        self.string_listbox.config(yscrollcommand=string_vscroll.set, xscrollcommand=string_hscroll.set)

        for col in remaining_string_cols:
            self.string_listbox.insert(tk.END, col)

        # Measures (numeric columns) label + listbox with scrollbars
        tk.Label(frame, text="Select Numeric Columns (Measures):").grid(row=0, column=1, sticky='w')

        numeric_frame = tk.Frame(frame)
        numeric_frame.grid(row=1, column=1, sticky='nsew', pady=5)
        self.numeric_listbox = tk.Listbox(numeric_frame, selectmode='multiple', exportselection=False, height=15)
        self.numeric_listbox.pack(side='left', fill='both', expand=True)

        numeric_vscroll = ttk.Scrollbar(numeric_frame, orient='vertical', command=self.numeric_listbox.yview)
        numeric_vscroll.pack(side='right', fill='y')
        numeric_hscroll = ttk.Scrollbar(frame, orient='horizontal', command=self.numeric_listbox.xview)
        numeric_hscroll.grid(row=2, column=1, sticky='ew')

        self.numeric_listbox.config(yscrollcommand=numeric_vscroll.set, xscrollcommand=numeric_hscroll.set)

        for col in self.numeric_cols:
            self.numeric_listbox.insert(tk.END, col)

        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=15)

        tk.Button(btn_frame, text="Submit", command=self.generate_table).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Restart Inspection", command=self.init_file_selection_page).pack(side='left', padx=5)

    def generate_table(self):
        newly_selected_strings = [self.string_listbox.get(i) for i in self.string_listbox.curselection()]
        selected_numeric_cols = [self.numeric_listbox.get(i) for i in self.numeric_listbox.curselection()]

        if not newly_selected_strings and not self.all_selected_string_cols:
            messagebox.showerror("Error", "Please select at least one dimension (string column).")
            return
        if not selected_numeric_cols:
            messagebox.showerror("Error", "Please select at least one measure (numeric column).")
            return

        # Append newly selected string columns (preserving order, avoid duplicates)
        for col in newly_selected_strings:
            if col not in self.all_selected_string_cols:
                self.all_selected_string_cols.append(col)

        self.selected_numeric_cols = selected_numeric_cols

        # Aggregate data
        orig_agg = self.df_original.groupby(self.all_selected_string_cols)[self.selected_numeric_cols].sum().reset_index()
        trans_agg = self.df_transformed.groupby(self.all_selected_string_cols)[self.selected_numeric_cols].sum().reset_index()

        merged = pd.merge(
            orig_agg,
            trans_agg,
            on=self.all_selected_string_cols,
            suffixes=('_orig', '_trans'),
            how='outer'
        )

        # Create concatenated filter key based on ALL selected string cols so far
        merged['_filter_key'] = merged[self.all_selected_string_cols].astype(str).agg(' | '.join, axis=1)

        # Detect anomaly for each numeric column
        for col in self.selected_numeric_cols:
            col_orig = f"{col}_orig"
            col_trans = f"{col}_trans"
            merged[col_orig] = pd.to_numeric(merged[col_orig], errors='coerce')
            merged[col_trans] = pd.to_numeric(merged[col_trans], errors='coerce')

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

        self.show_table(merged, "_".join(self.all_selected_string_cols))
        self.export_to_excel(merged, "_".join(self.all_selected_string_cols))

    def show_table(self, df, combo_name):
        self.clear_root()
        frame = tk.Frame(self.root)
        frame.pack(fill='both', expand=True)

        tk.Label(frame, text=f"Reconciliation Results ({combo_name})").pack(pady=5)

        table_frame = tk.Frame(frame)
        table_frame.pack(fill='both', expand=True)

        # Treeview with scrollbars
        columns = list(df.columns)
        tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        tree.pack(side='left', fill='both', expand=True)

        vscroll = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        vscroll.pack(side='right', fill='y')
        hscroll = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
        hscroll.pack(fill='x')

        tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor='w')

        for _, row in df.iterrows():
            tree.insert('', 'end', values=list(row))

        self.filter_listbox = None

        # Filter UI (multi-select) for _filter_key values
        filter_frame = tk.Frame(frame)
        filter_frame.pack(fill='x', pady=10)

        tk.Label(filter_frame, text="Filter Rows by concatenated key:").pack(anchor='w')

        self.filter_listbox = tk.Listbox(filter_frame, selectmode='multiple', height=8, exportselection=False)
        self.filter_listbox.pack(side='left', fill='x', expand=True)

        # Add scrollbar for filter listbox
        filter_scrollbar = ttk.Scrollbar(filter_frame, orient='vertical', command=self.filter_listbox.yview)
        filter_scrollbar.pack(side='right', fill='y')
        self.filter_listbox.config(yscrollcommand=filter_scrollbar.set)

        # Populate filter listbox with unique _filter_key values in original order
        filter_keys = df['_filter_key'].dropna().unique()
        for val in filter_keys:
            self.filter_listbox.insert(tk.END, val)

        # Buttons below filter listbox
        btn_frame = tk.Frame(frame)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Drill Down Next", command=lambda: self.drill_down(df)).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Restart Inspection", command=self.init_file_selection_page).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Export Current Table to Excel", command=lambda: self.export_to_excel(df, combo_name)).pack(side='left', padx=5)

    def drill_down(self, df):
        if self.filter_listbox is None:
            messagebox.showerror("Error", "Filter list not available.")
            return

        selected_indices = self.filter_listbox.curselection()
        if not selected_indices:
            # No filter selection, use all rows
            filtered_df = df.copy()
        else:
            selected_keys = [self.filter_listbox.get(i) for i in selected_indices]
            filtered_df = df[df['_filter_key'].isin(selected_keys)]

        # Now filter the full original and transformed dfs to only those matching selected keys in ALL selected string columns
        filter_mask_orig = pd.Series(True, index=self.df_original_full.index)
        filter_mask_trans = pd.Series(True, index=self.df_transformed_full.index)

        for col in self.all_selected_string_cols:
            vals = filtered_df[col].dropna().unique().astype(str)
            filter_mask_orig &= self.df_original_full[col].astype(str).isin(vals)
            filter_mask_trans &= self.df_transformed_full[col].astype(str).isin(vals)

        self.df_original = self.df_original_full[filter_mask_orig].copy()
        self.df_transformed = self.df_transformed_full[filter_mask_trans].copy()

        # Move to next column selection page to select additional columns (if any)
        self.init_column_selection_page()

    def export_to_excel(self, df, combo_name):
        # Setup workbook path and open/create workbook
        if self.export_wb_path is None:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Reconcile_{timestamp}.xlsx"
            self.export_wb_path = os.path.join(self.output_folder if self.output_folder else os.getcwd(), filename)
            self.export_wb = Workbook()
            default_sheet = self.export_wb.active
            self.export_wb.remove(default_sheet)
            self.export_counter = 1
        else:
            try:
                self.export_wb = load_workbook(self.export_wb_path)
                self.export_counter += 1
            except Exception:
                # If loading fails, create new workbook (rare)
                self.export_wb = Workbook()
                default_sheet = self.export_wb.active
                self.export_wb.remove(default_sheet)
                self.export_counter = 1

        sheet_data_name = f"{combo_name}{self.export_counter}"
        sheet_chart_name = f"{combo_name}{self.export_counter}chart"

        # Add data sheet
        ws = self.export_wb.create_sheet(title=sheet_data_name)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Add chart sheet
        chart_ws = self.export_wb.create_sheet(title=sheet_chart_name)

        for col in self.selected_numeric_cols:
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            df_chart = df[['_filter_key', f"{col}_orig", f"{col}_trans"]].copy()
            df_chart.set_index('_filter_key', inplace=True)

            plt.figure(figsize=(8, 5))
            df_chart[[f"{col}_orig", f"{col}_trans"]].plot(kind='bar')
            plt.title(f"Original vs Transformed - {col}")
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            plt.savefig(temp_file.name)
            plt.close()

            img = Image(temp_file.name)
            next_row = chart_ws.max_row + 2 if chart_ws.max_row > 1 else 1
            chart_ws.add_image(img, f"A{next_row}")

            temp_file.close()
            os.unlink(temp_file.name)

        try:
            self.export_wb.save(self.export_wb_path)
            messagebox.showinfo("Exported", f"Workbook saved/appended:\n{self.export_wb_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel workbook:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1100x750")
    app = DataReconciliationApp(root)
    root.mainloop()

```
