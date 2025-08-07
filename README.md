
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
    
