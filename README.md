import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

class ReconciliationApp:
    def __init__(self, root, df_original, df_transformed):
        self.root = root
        self.df_original = df_original.copy()
        self.df_transformed = df_transformed.copy()
        self.all_string_cols = [col for col in df_original.columns if df_original[col].dtype == 'object']
        self.all_numeric_cols = [col for col in df_original.columns if pd.api.types.is_numeric_dtype(df_original[col])]
        self.selected_string_cols = []
        self.selected_numeric_cols = []
        self.string_vars = []
        self.numeric_vars = []

        self.setup_selection_gui()

    def clear_gui(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def setup_selection_gui(self):
        self.clear_gui()

        tk.Label(self.root, text="Select String Columns for Grouping:").pack(pady=(10, 0))

        self.string_vars = []
        for col in self.all_string_cols:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(self.root, text=col, variable=var)
            chk.pack(anchor="w")
            self.string_vars.append((col, var))

        tk.Label(self.root, text="Select Numeric Columns to Aggregate:").pack(pady=(10, 0))

        self.numeric_vars = []
        for col in self.all_numeric_cols:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(self.root, text=col, variable=var)
            chk.pack(anchor="w")
            self.numeric_vars.append((col, var))

        tk.Button(self.root, text="Submit & Show Aggregated Table", command=self.show_aggregated_table).pack(pady=10)  ## new 3

    def show_aggregated_table(self):  ## new 3
        self.selected_string_cols = [col for col, var in self.string_vars if var.get()]
        self.selected_numeric_cols = [col for col, var in self.numeric_vars if var.get()]

        if not self.selected_string_cols or not self.selected_numeric_cols:
            messagebox.showerror("Selection Error", "Select both string and numeric columns.")
            return

        def concat_key(df):
            return df[self.selected_string_cols].astype(str).agg(' | '.join, axis=1)

        df_o = self.df_original.copy()
        df_t = self.df_transformed.copy()

        df_o["Filter_Key"] = concat_key(df_o)
        df_t["Filter_Key"] = concat_key(df_t)

        df_o_grouped = df_o.groupby("Filter_Key")[self.selected_numeric_cols].sum().reset_index()
        df_t_grouped = df_t.groupby("Filter_Key")[self.selected_numeric_cols].sum().reset_index()

        merged = pd.merge(
            df_o_grouped,
            df_t_grouped,
            on="Filter_Key",
            how="outer",
            suffixes=('_original', '_transformed')
        )

        for col in self.selected_numeric_cols:
            merged[f"{col}_diff"] = merged[f"{col}_original"].fillna(0) - merged[f"{col}_transformed"].fillna(0)

        self.display_table(merged)

    def display_table(self, df):  ## new 3
        self.clear_gui()
        tk.Label(self.root, text="Aggregated Comparison Table").pack(pady=5)

        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, columns=list(df.columns), show="headings")
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="w")

        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

        tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Back", command=self.setup_selection_gui).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="Export to Excel", command=lambda: self.export_to_excel(df)).grid(row=0, column=1, padx=5)

    def export_to_excel(self, df):  ## new 3
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Exported", f"Aggregated data saved to:\n{file_path}")


# Test run
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
    root.title("Grouped Aggregation Reconciliation Tool")
    root.geometry("1200x600")
    app = ReconciliationApp(root, df_original, df_transformed)
    root.mainloop()








    
        
        
        
        
        
        tk.Button(btn_frame, text="Run Final Reconciliation", command=self.reconcile).grid(row=0, column=1, padx=5)
        tk.Button(btn_frame, text="Back to Column Selection", command=self.setup_selection_gui).grid(row=0, column=2, padx=5)

    def drill_down(self):  ## new 2
        if not self.remaining_string_cols:
            messagebox.showinfo("No More Columns", "No more string columns left to drill down.")
            return
        self.setup_selection_gui()

    def reconcile(self):
        if not self.used_string_cols or not self.selected_numeric_cols:  ## new 2
            messagebox.showerror("Selection Error", "Ensure at least one string column and numeric column are selected.")  ## new 2
            return

        original_grouped = self.df_original.groupby(self.used_string_cols)[self.selected_numeric_cols].sum().reset_index()  ## new 2
        transformed_grouped = self.df_transformed.groupby(self.used_string_cols)[self.selected_numeric_cols].sum().reset_index()  ## new 2

        merged = pd.merge(
            original_grouped,
            transformed_grouped,
            on=self.used_string_cols,
            how='outer',
            suffixes=('_original', '_transformed')
        )

        for col in self.selected_numeric_cols:  ## new 2
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
    app = ReconciliationApp(root, df_original, df_transformed)
    root.mainloop()

<!---
AditiKollur/AditiKollur is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
