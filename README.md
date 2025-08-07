
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

class ReconciliationApp:
    def __init__(self, root, df_original, df_transformed):
        self.root = root
        self.df_original_full = df_original.copy()
        self.df_transformed_full = df_transformed.copy()
        self.df_original = df_original.copy()
        self.df_transformed = df_transformed.copy()

        self.used_string_cols = []
        self.remaining_string_cols = [col for col in df_original.columns if df_original[col].dtype == 'object']
        self.numeric_cols = [col for col in df_original.columns if pd.api.types.is_numeric_dtype(df_original[col])]
        self.selected_numeric_cols = []
        self.string_vars = []

        self.setup_selection_gui()

    def clear_gui(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def setup_selection_gui(self):
        self.clear_gui()

        tk.Label(self.root, text="Select String Columns for Matching:").pack(pady=(10, 0))

        self.string_vars = []
        for col in self.remaining_string_cols:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(self.root, text=col, variable=var)
            chk.pack(anchor="w")
            self.string_vars.append((col, var))

        tk.Label(self.root, text="Select Numeric Columns to Reconcile:").pack(pady=(10, 0))

        self.numeric_vars = []
        for col in self.numeric_cols:
            var = tk.BooleanVar(value=(col in self.selected_numeric_cols))
            chk = tk.Checkbutton(self.root, text=col, variable=var)
            chk.pack(anchor="w")
            self.numeric_vars.append((col, var))

        tk.Button(self.root, text="Submit & Show Filtered Table", command=self.apply_filter).pack(pady=10)  ## new 2

    def apply_filter(self):  ## new 2
        new_keys = [col for col, var in self.string_vars if var.get()]
        if not new_keys:
            messagebox.showwarning("Warning", "Select at least one column to filter.")
            return

        self.used_string_cols.extend(new_keys)
        self.remaining_string_cols = [col for col in self.remaining_string_cols if col not in new_keys]

        self.selected_numeric_cols = [col for col, var in self.numeric_vars if var.get()]
        if not self.selected_numeric_cols:
            messagebox.showerror("Selection Error", "Please select numeric columns to compare.")
            return

        def concat_key(df):
            return df[self.used_string_cols].astype(str).agg(' | '.join, axis=1)

        self.df_original["__filter_key__"] = concat_key(self.df_original)
        self.df_transformed["__filter_key__"] = concat_key(self.df_transformed)

        common_keys = set(self.df_original["__filter_key__"]).intersection(set(self.df_transformed["__filter_key__"]))
        self.df_original = self.df_original[self.df_original["__filter_key__"].isin(common_keys)].copy()
        self.df_transformed = self.df_transformed[self.df_transformed["__filter_key__"].isin(common_keys)].copy()

        self.display_filtered_table()

    def display_filtered_table(self):  ## new 2
        self.clear_gui()

        tk.Label(self.root, text=f"Filtered Data Preview (Rows: {len(self.df_original)})").pack(pady=5)

        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, columns=list(self.df_original.columns), show="headings")
        for col in self.df_original.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="w")

        tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        for _, row in self.df_original.iterrows():
            tree.insert("", "end", values=list(row))

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Next (Select More Columns)", command=self.setup_selection_gui).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="Export Filtered to Excel", command=self.export_filtered).grid(row=0, column=1, padx=5)
        tk.Button(btn_frame, text="Run Final Reconciliation", command=self.reconcile).grid(row=0, column=2, padx=5)

    def export_filtered(self):  ## new 2
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                self.df_original.to_excel(writer, index=False, sheet_name='Filtered_Original')
                self.df_transformed.to_excel(writer, index=False, sheet_name='Filtered_Transformed')
            messagebox.showinfo("Exported", f"Filtered data saved to:\n{file_path}")

    def reconcile(self):
        if not self.used_string_cols or not self.selected_numeric_cols:
            messagebox.showerror("Selection Error", "Ensure at least one string column and numeric column are selected.")
            return

        original_grouped = self.df_original.groupby(self.used_string_cols)[self.selected_numeric_cols].sum().reset_index()
        transformed_grouped = self.df_transformed.groupby(self.used_string_cols)[self.selected_numeric_cols].sum().reset_index()

        merged = pd.merge(
            original_grouped,
            transformed_grouped,
            on=self.used_string_cols,
            how='outer',
            suffixes=('_original', '_transformed')
        )

        for col in self.selected_numeric_cols:
            col_orig = f"{col}_original"
            col_trns = f"{col}_transformed"
            merged[col_orig] = merged[col_orig].astype(float)
            merged[col_trns] = merged[col_trns].astype(float)

            def get_anomaly(row):
                v1, v2 = row[col_orig], row[col_trns]
                if pd.isna(v1) or pd.isna(v2):
                    return "anomaly"
                return "OK" if v1 == v2 else "anomaly"

            merged[f"{col}_anomaly"] = merged.apply(get_anomaly, axis=1)
            merged[col_orig] = merged[col_orig].fillna("Missing")
            merged[col_trns] = merged[col_trns].fillna("Missing")

        self.display_results(merged)

    def display_results(self, df_result):
        self.clear_gui()
        tk.Label(self.root, text="Reconciliation Results").pack()

        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame)
        tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scrollbar.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scrollbar.set)

        tree["columns"] = list(df_result.columns)
        tree["show"] = "headings"

        for col in df_result.columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="w", width=120)

        for _, row in df_result.iterrows():
            tree.insert("", "end", values=list(row))

        tk.Button(self.root, text="Back", command=self.setup_selection_gui).pack(pady=10)


# Example test
if __name__ == "__main__":
    df_original = pd.DataFrame({
        'Region': ['North', 'South', 'East', 'North'],
        'Product': ['A', 'B', 'C', 'D'],
        'Category': ['X', 'Y', 'Z', 'X'],
        'Sales': [100, 200, 150, 300],
        'Profit': [10, 20, 15, 25]
    })

    df_transformed = pd.DataFrame({
        'Region': ['North', 'South', 'West', 'North'],
        'Product': ['A', 'B', 'C', 'D'],
        'Category': ['X', 'Y', 'Z', 'X'],
        'Sales': [100, 250, 150, 280],
        'Profit': [10, 22, 15, 27]
    })

    root = tk.Tk()
    root.title("Multi-step Reconciliation Tool")
    root.geometry("1200x600")  ## new 2
    app = ReconciliationApp(root, df_original, df_transformed)
    root.mainloop()
