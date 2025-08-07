import tkinter as tk from tkinter import ttk, messagebox, filedialog import pandas as pd from openpyxl import Workbook from openpyxl.chart import BarChart, Reference from openpyxl.utils.dataframe import dataframe_to_rows

class ReconciliationApp: def init(self, root, df_original, df_transformed): self.root = root self.df_original = df_original.copy() self.df_transformed = df_transformed.copy() self.all_string_columns = [col for col in df_original.columns if df_original[col].dtype == 'object'] self.numeric_columns = [col for col in df_original.columns if pd.api.types.is_numeric_dtype(df_original[col])] self.selected_keys = [] self.result_df = None self.setup_selection_gui()

def clear_gui(self):
    for widget in self.root.winfo_children():
        widget.destroy()

def setup_selection_gui(self):
    self.clear_gui()
    tk.Label(self.root, text="Select String Columns (keys for match):").pack()

    self.string_vars = []
    for col in self.all_string_columns:
        if col not in self.selected_keys:
            var = tk.BooleanVar()
            tk.Checkbutton(self.root, text=col, variable=var).pack(anchor="w")
            self.string_vars.append((col, var))

    tk.Label(self.root, text="Select Numeric Columns (to compare):").pack(pady=(10, 0))

    self.numeric_vars = []
    for col in self.numeric_columns:
        var = tk.BooleanVar()
        tk.Checkbutton(self.root, text=col, variable=var).pack(anchor="w")
        self.numeric_vars.append((col, var))

    tk.Button(self.root, text="Submit", command=self.reconcile).pack(pady=10)

def reconcile(self):
    new_keys = [col for col, var in self.string_vars if var.get()]
    numeric_cols = [col for col, var in self.numeric_vars if var.get()]

    if not new_keys or not numeric_cols:
        messagebox.showerror("Selection Error", "Please select both key and numeric columns.")
        return

    self.selected_keys.extend(new_keys)

    key_cols = self.selected_keys

    # Group and aggregate numeric values
    original_grouped = self.df_original.groupby(key_cols)[numeric_cols].sum().reset_index()
    transformed_grouped = self.df_transformed.groupby(key_cols)[numeric_cols].sum().reset_index()

    merged = pd.merge(
        original_grouped,
        transformed_grouped,
        on=key_cols,
        how='outer',
        suffixes=('_original', '_transformed')
    )

    for col in numeric_cols:
        col_orig = f"{col}_original"
        col_trns = f"{col}_transformed"

        merged[col_orig] = pd.to_numeric(merged[col_orig], errors='coerce')
        merged[col_trns] = pd.to_numeric(merged[col_trns], errors='coerce')

        def get_anomaly(row):
            v1, v2 = row[col_orig], row[col_trns]
            if pd.isna(v1) or pd.isna(v2):
                return "missing"
            elif v1 != v2:
                return "anomaly"
            return "ok"

        merged[f"{col}_anomaly"] = merged.apply(get_anomaly, axis=1)
        merged[col_orig] = merged[col_orig].fillna("Missing")
        merged[col_trns] = merged[col_trns].fillna("Missing")

    merged['_filter_key'] = merged[key_cols].astype(str).agg('|'.join, axis=1)
    self.result_df = merged
    self.display_results(merged)

def display_results(self, df_result):
    self.clear_gui()

    tk.Label(self.root, text="Reconciliation Results").pack()

    filter_values = sorted(df_result['_filter_key'].unique())
    self.filter_var = tk.StringVar(value=filter_values)

    filter_label = tk.Label(self.root, text="Filter by Key:")
    filter_label.pack()

    filter_box = tk.Listbox(self.root, listvariable=self.filter_var, selectmode="multiple", height=5)
    filter_box.pack()
    self.filter_box = filter_box

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

    self.tree = tree

    tk.Button(self.root, text="Next Drill Down", command=self.setup_selection_gui).pack(pady=5)
    tk.Button(self.root, text="Export to Excel", command=self.export_to_excel).pack(pady=5)

def export_to_excel(self):
    if self.result_df is None:
        messagebox.showerror("No Data", "No data to export")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    df = self.result_df.copy()
    wb = Workbook()
    ws_data = wb.active
    ws_data.title = "Reconciliation Data"

    for r in dataframe_to_rows(df, index=False, header=True):
        ws_data.append(r)

    ws_chart = wb.create_sheet(title="Bar Chart")

    chart_data = df.melt(id_vars=['_filter_key'], var_name='Metric', value_name='Value')
    pivot_df = chart_data.pivot(index='_filter_key', columns='Metric', values='Value').reset_index()

    for r in dataframe_to_rows(pivot_df, index=False, header=True):
        ws_chart.append(r)

    chart = BarChart()
    chart.type = "col"
    chart.title = "Original vs Transformed Metrics"
    chart.y_axis.title = "Value"
    chart.x_axis.title = "_filter_key"

    num_cols = len(pivot_df.columns)
    data = Reference(ws_chart, min_col=2, max_col=num_cols, min_row=1, max_row=ws_chart.max_row)
    cats = Reference(ws_chart, min_col=1, min_row=2, max_row=ws_chart.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    ws_chart.add_chart(chart, "H2")
    wb.save(file_path)
    messagebox.showinfo("Exported", f"Data exported to {file_path}")

Example use

if name == "main": df_original = pd.DataFrame({ 'Region': ['North', 'South', 'East'], 'Product': ['A', 'B', 'C'], 'Sales': [100, 200, 150], 'Profit': [10, 20, 15] })

df_transformed = pd.DataFrame({
    'Region': ['North', 'South', 'East'],
    'Product': ['A', 'B', 'C'],
    'Sales': [100, 250, 150],
    'Profit': [10, 22, 15]
})

root = tk.Tk()
root.title("Multi-Step Reconciliation Tool")
app = ReconciliationApp(root, df_original, df_transformed)
root.mainloop()

