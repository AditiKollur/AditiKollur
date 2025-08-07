import tkinter as tk from tkinter import ttk, messagebox, filedialog import pandas as pd import os from openpyxl import Workbook from openpyxl.utils.dataframe import dataframe_to_rows from openpyxl.chart import BarChart, Reference from openpyxl.styles import Font

class ReconciliationApp: def init(self, root, df_original, df_transformed): self.root = root self.df_original = df_original self.df_transformed = df_transformed self.selected_keys = [] self.remaining_keys = [] self.selected_values = None  # ##new self.numeric_cols = [] self.drill_level = 0

self.setup_selection_gui()

def clear_gui(self):
    for widget in self.root.winfo_children():
        widget.destroy()

def setup_selection_gui(self):
    self.clear_gui()

    tk.Label(self.root, text=f"Select String Columns for Drill Down Level {self.drill_level+1}:").pack()
    self.string_vars = []
    current_cols = [col for col in self.df_original.columns if self.df_original[col].dtype == 'object']
    current_cols = [col for col in current_cols if col not in self.selected_keys]

    for col in current_cols:
        var = tk.BooleanVar()
        tk.Checkbutton(self.root, text=col, variable=var).pack(anchor="w")
        self.string_vars.append((col, var))

    tk.Label(self.root, text="Select Numeric Columns (to compare):").pack(pady=(10, 0))
    self.numeric_vars = []
    for col in self.df_original.columns:
        if pd.api.types.is_numeric_dtype(self.df_original[col]):
            var = tk.BooleanVar()
            tk.Checkbutton(self.root, text=col, variable=var).pack(anchor="w")
            self.numeric_vars.append((col, var))

    tk.Button(self.root, text="Submit", command=self.reconcile).pack(pady=10)

def reconcile(self):
    new_keys = [col for col, var in self.string_vars if var.get()]
    self.numeric_cols = [col for col, var in self.numeric_vars if var.get()]

    if not new_keys or not self.numeric_cols:
        messagebox.showerror("Selection Error", "Please select both key and numeric columns.")
        return

    self.selected_keys.extend(new_keys)

    df1 = self.df_original.copy()
    df2 = self.df_transformed.copy()

    if self.selected_values is not None:
        df1 = df1[df1['_filter_key'].isin(self.selected_values)]
        df2 = df2[df2['_filter_key'].isin(self.selected_values)]

    df1['_filter_key'] = df1[self.selected_keys].astype(str).agg(' | '.join, axis=1)
    df2['_filter_key'] = df2[self.selected_keys].astype(str).agg(' | '.join, axis=1)

    grouped1 = df1.groupby(['_filter_key'] + self.selected_keys)[self.numeric_cols].sum().reset_index()
    grouped2 = df2.groupby(['_filter_key'] + self.selected_keys)[self.numeric_cols].sum().reset_index()

    merged = pd.merge(
        grouped1,
        grouped2,
        on=['_filter_key'] + self.selected_keys,
        how='outer',
        suffixes=('_original', '_transformed')
    )

    for col in self.numeric_cols:
        col_orig = f"{col}_original"
        col_trns = f"{col}_transformed"

        merged[col_orig] = merged[col_orig].astype('float64')
        merged[col_trns] = merged[col_trns].astype('float64')

        def get_anomaly(row):
            v1, v2 = row[col_orig], row[col_trns]
            if pd.isna(v1) or pd.isna(v2):
                return "Missing"
            elif v1 != v2:
                return "Anomaly"
            return "OK"

        merged[f"{col}_anomaly"] = merged.apply(get_anomaly, axis=1)
        merged[col_orig] = merged[col_orig].fillna("Missing")
        merged[col_trns] = merged[col_trns].fillna("Missing")

    self.display_results(merged)

def display_results(self, df_result):
    self.clear_gui()
    self.df_result = df_result

    tk.Label(self.root, text="Reconciliation Results").pack()

    frame_top = tk.Frame(self.root)
    frame_top.pack(fill="x")

    tk.Label(frame_top, text="Filter by _filter_key:").pack(side="left")
    self.filter_var = tk.StringVar(value=df_result['_filter_key'].unique().tolist())

    self.filter_box = tk.Listbox(frame_top, listvariable=self.filter_var, selectmode="multiple", height=5, width=80)
    self.filter_box.pack(side="left", padx=5)

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

    tk.Button(self.root, text="Export to Excel", command=self.export_to_excel).pack(pady=5)
    tk.Button(self.root, text="Next Drill Down", command=self.next_drill_down).pack(pady=5)
    tk.Button(self.root, text="Back", command=self.setup_selection_gui).pack(pady=5)

def next_drill_down(self):
    selected = [self.filter_box.get(i) for i in self.filter_box.curselection()]
    self.selected_values = selected if selected else None
    self.drill_level += 1
    self.setup_selection_gui()

def export_to_excel(self):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Reconciliation Data"

    for r in dataframe_to_rows(self.df_result, index=False, header=True):
        ws1.append(r)

    # Create bar charts for each selected key
    chart_sheet = wb.create_sheet(title="Bar Chart")
    for i, key in enumerate(self.selected_keys):
        for col in self.numeric_cols:
            col_orig = f"{col}_original"
            col_trns = f"{col}_transformed"

            filtered_df = self.df_result[[key, '_filter_key', col_orig, col_trns]].dropna()
            chart_data = filtered_df.groupby(key)[[col_orig, col_trns]].sum().reset_index()

            start_row = i * (len(chart_data) + 8) + 1

            for row_idx, row in enumerate(dataframe_to_rows(chart_data, index=False, header=True), start=start_row):
                chart_sheet.append(row)

            chart = BarChart()
            chart.title = f"{col} by {key}"
            chart.y_axis.title = col
            chart.x_axis.title = key

            data_ref = Reference(chart_sheet, min_col=2, max_col=3,
                                 min_row=start_row, max_row=start_row + len(chart_data))
            cats_ref = Reference(chart_sheet, min_col=1, min_row=start_row + 1,
                                 max_row=start_row + len(chart_data))

            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)

            chart_sheet.add_chart(chart, f"E{start_row}")

    wb.save(file_path)

Sample usage

if name == "main": df_original = pd.DataFrame({ 'Region': ['North', 'South', 'East', 'North'], 'Product': ['A', 'B', 'C', 'A'], 'Sales': [100, 200, 150, 120], 'Profit': [10, 20, 15, 11] })

df_transformed = pd.DataFrame({
    'Region': ['North', 'South', 'East', 'West'],
    'Product': ['A', 'B', 'C', 'C'],
    'Sales': [100, 250, 150, 50],
    'Profit': [10, 22, 15, 5]
})

root = tk.Tk()
root.title("Reconciliation Tool")
app = ReconciliationApp(root, df_original, df_transformed)
root.mainloop()

