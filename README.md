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
        self.df_original_full = None   # store full original df (never lose data)
        self.df_transformed_full = None
        self.df_original = None        # filtered copies used in drill down
        self.df_transformed = None
        self.string_cols = []
        self.numeric_cols = []
        self.selected_string_cols = []
        self.selected_numeric_cols = []
        self.selected_values = []
        self.export_wb_path = None  # Path of the export workbook (created on first export)
        self.export_wb = None       # Loaded Workbook object for appending sheets
        self.export_counter = 0     # To number sheets for each export
        self.init_file_selection_page()

    def clear_root(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def drop_single_unique_columns(self, df):
        # Drop columns with only 1 unique value including NaN
        return df.loc[:, df.nunique(dropna=False) > 1]

    def init_file_selection_page(self):
        self.clear_root()
        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text="Select Original File:").pack(pady=(10,0))
        tk.Button(frame, text="Browse", command=self.load_original_file).pack(pady=(0,10))
        tk.Label(frame, text="Select Transformed File:").pack(pady=(10,0))
        tk.Button(frame, text="Browse", command=self.load_transformed_file).pack(pady=(0,10))
        tk.Button(frame, text="Next", command=self.init_column_selection_page).pack(pady=20)

        # Reset all state on returning to first page
        self.df_original_full = None
        self.df_transformed_full = None
        self.df_original = None
        self.df_transformed = None
        self.export_wb_path = None
        self.export_wb = None
        self.export_counter = 0

    def load_original_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.df_original_full = pd.read_excel(file_path)
            messagebox.showinfo("Loaded", "Original file loaded successfully!")

    def load_transformed_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.df_transformed_full = pd.read_excel(file_path)
            messagebox.showinfo("Loaded", "Transformed file loaded successfully!")

    def init_column_selection_page(self):
        if self.df_original_full is None or self.df_transformed_full is None:
            messagebox.showerror("Error", "Please load both files before proceeding.")
            return

        # Reset filtered dfs to full copies on new selection page
        self.df_original = self.df_original_full.copy()
        self.df_transformed = self.df_transformed_full.copy()

        # Drop columns with only 1 unique value including NaN from both DataFrames
        df_orig_trimmed = self.drop_single_unique_columns(self.df_original)
        df_trans_trimmed = self.drop_single_unique_columns(self.df_transformed)

        # Find common columns after dropping single-unique-value columns
        common_cols = list(set(df_orig_trimmed.columns) & set(df_trans_trimmed.columns))

        # Separate string and numeric columns from trimmed original DataFrame
        self.string_cols = [c for c in common_cols if df_orig_trimmed[c].dtype == 'object']
        self.numeric_cols = [c for c in common_cols if pd.api.types.is_numeric_dtype(df_orig_trimmed[c])]

        self.clear_root()
        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        # String columns selection
        tk.Label(frame, text="Select String Columns (Dimensions):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.string_listbox = tk.Listbox(frame, selectmode="multiple", exportselection=False, height=10)
        for col in self.string_cols:
            self.string_listbox.insert(tk.END, col)
        self.string_listbox.grid(row=1, column=0, sticky="nsew", padx=(10,0), pady=5)
        scrollbar1 = tk.Scrollbar(frame, command=self.string_listbox.yview)
        scrollbar1.grid(row=1, column=1, sticky="ns", pady=5)
        self.string_listbox.config(yscrollcommand=scrollbar1.set)

        # Numeric columns selection
        tk.Label(frame, text="Select Numeric Columns (Measures):").grid(row=0, column=2, padx=10, pady=5, sticky="w")
        self.numeric_listbox = tk.Listbox(frame, selectmode="multiple", exportselection=False, height=10)
        for col in self.numeric_cols:
            self.numeric_listbox.insert(tk.END, col)
        self.numeric_listbox.grid(row=1, column=2, sticky="nsew", padx=(10,0), pady=5)
        scrollbar2 = tk.Scrollbar(frame, command=self.numeric_listbox.yview)
        scrollbar2.grid(row=1, column=3, sticky="ns", pady=5)
        self.numeric_listbox.config(yscrollcommand=scrollbar2.set)

        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=15)

        tk.Button(btn_frame, text="Submit", command=self.generate_table).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Restart Inspection", command=self.init_file_selection_page).pack(side='left', padx=5)  # Refresh button

    def generate_table(self):
        self.selected_string_cols = [self.string_cols[i] for i in self.string_listbox.curselection()]
        self.selected_numeric_cols = [self.numeric_cols[i] for i in self.numeric_listbox.curselection()]

        if not self.selected_string_cols or not self.selected_numeric_cols:
            messagebox.showerror("Error", "Please select at least one dimension and one measure.")
            return

        # Aggregate original and transformed DataFrames on selected columns
        orig_agg = self.df_original.groupby(self.selected_string_cols)[self.selected_numeric_cols].sum().reset_index()
        trans_agg = self.df_transformed.groupby(self.selected_string_cols)[self.selected_numeric_cols].sum().reset_index()

        # Merge aggregated results
        merged = pd.merge(
            orig_agg,
            trans_agg,
            on=self.selected_string_cols,
            suffixes=('_orig', '_trans'),
            how='outer'
        )

        # Create concatenated filter key from selected string columns (for filtering in UI)
        combo_name = "_".join(self.selected_string_cols)
        merged['_filter_key'] = merged[self.selected_string_cols].astype(str).agg(' | '.join, axis=1)

        # Add anomaly status columns for each numeric measure
        for col in self.selected_numeric_cols:
            status_col = f"{col}_status"
            def anomaly_status(row):
                o = row.get(f"{col}_orig")
                t = row.get(f"{col}_trans")
                if pd.isna(o) or pd.isna(t):
                    return "Missing"
                elif abs(o - t) < 1e-9:
                    return "OK"
                else:
                    return "Anomaly"
            merged[status_col] = merged.apply(anomaly_status, axis=1)

        self.show_table(merged, combo_name)

        # Export to Excel, appending sheets to existing workbook or creating new one
        self.export_to_excel(merged, combo_name)

    def show_table(self, df, combo_name):
        self.clear_root()
        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        # Filter multi-select listbox
        tk.Label(frame, text="Filter _filter_key (multiple select):").pack(anchor="w", padx=10, pady=(10,0))

        self.filter_listbox = tk.Listbox(frame, selectmode='multiple', height=8)
        self.filter_listbox.pack(fill='x', padx=10)
        scrollbar_filter = tk.Scrollbar(frame, orient='vertical', command=self.filter_listbox.yview)
        scrollbar_filter.pack(side='right', fill='y', padx=(0,10), pady=(0,130))
        self.filter_listbox.config(yscrollcommand=scrollbar_filter.set)

        for val in df['_filter_key'].unique():
            self.filter_listbox.insert(tk.END, val)

        # Treeview with scrollbars
        tree_frame = tk.Frame(frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = list(df.columns)
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        tree.pack(side='left', fill='both', expand=True)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor='w')

        for _, row in df.iterrows():
            tree.insert('', tk.END, values=list(row))

        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        vsb.pack(side='right', fill='y')
        hsb = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
        hsb.pack(fill='x', padx=10)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        btn_frame = tk.Frame(frame)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Filter & Drill Down", command=lambda: self.drill_down(df)).pack(side='left', padx=10)
        tk.Button(btn_frame, text="Restart Inspection", command=self.init_file_selection_page).pack(side='left', padx=10)

    def drill_down(self, df):
        selected_indices = self.filter_listbox.curselection()
        if not selected_indices:
            filtered_df = df.copy()
        else:
            selected_keys = [self.filter_listbox.get(i) for i in selected_indices]
            filtered_df = df[df['_filter_key'].isin(selected_keys)]

        keys_split = filtered_df['_filter_key'].str.split(' \| ', expand=True)

        # Filter on original full dfs (do NOT lose unfiltered data)
        filter_mask_orig = pd.Series(False, index=self.df_original_full.index)
        filter_mask_trans = pd.Series(False, index=self.df_transformed_full.index)
        for i, col in enumerate(self.selected_string_cols):
            vals = keys_split[i].unique()
            filter_mask_orig |= self.df_original_full[col].astype(str).isin(vals)
            filter_mask_trans |= self.df_transformed_full[col].astype(str).isin(vals)

        self.df_original = self.df_original_full[filter_mask_orig].copy()
        self.df_transformed = self.df_transformed_full[filter_mask_trans].copy()

        self.init_column_selection_page()

    def export_to_excel(self, df, combo_name):
        # Create filename once with timestamp if not already created
        if self.export_wb_path is None:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Reconcile_{timestamp}.xlsx"
            self.export_wb_path = os.path.join(os.getcwd(), filename)
            self.export_wb = Workbook()
            # Remove default sheet created by openpyxl
            default_sheet = self.export_wb.active
            self.export_wb.remove(default_sheet)
            self.export_counter = 1
        else:
            # Load existing workbook
            self.export_wb = load_workbook(self.export_wb_path)
            self.export_counter += 1

        # Sheet names
        sheet_data_name = f"{combo_name}{self.export_counter}"
        sheet_chart_name = f"{combo_name}{self.export_counter}chart"

        # Add reconciliation data sheet
        ws = self.export_wb.create_sheet(title=sheet_data_name)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Add charts sheet with bar chart for each numeric column
        chart_ws = self.export_wb.create_sheet(title=sheet_chart_name)

        for col in self.selected_numeric_cols:
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            df_chart = df[['_filter_key', f"{col}_orig", f"{col}_trans"]]

            plt.figure(figsize=(8, 5))
            df_chart.set_index('_filter_key')[[f"{col}_orig", f"{col}_trans"]].plot(kind='bar')
            plt.title(f"Original vs Transformed - {col}")
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            plt.savefig(temp_file.name)
            plt.close()

            img = Image(temp_file.name)
            # Find next free row for placing images vertically stacked
            next_row = chart_ws.max_row + 2 if chart_ws.max_row > 1 else 1
            chart_ws.add_image(img, f"A{next_row}")
            temp_file.close()
            os.unlink(temp_file.name)

        # Save workbook
        try:
            self.export_wb.save(self.export_wb_path)
            messagebox.showinfo("Exported", f"Workbook saved/appended:\n{self.export_wb_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel workbook:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1000x700")
    app = DataReconciliationApp(root)
    root.mainloop()

```
