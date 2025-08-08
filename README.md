```python
import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows


class ReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Reconciliation Tool")
        self.df_original = None
        self.df_transformed = None
        self.selected_values = None
        self.remaining_columns = []
        self.numeric_columns = []
        self.string_columns = []
        self.selected_string_columns = []

        self.page1()

    def page1(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        tk.Button(frame, text="Load Original File", command=self.load_original).pack(pady=5)
        tk.Button(frame, text="Load Transformed File", command=self.load_transformed).pack(pady=5)
        tk.Button(frame, text="Next", command=self.page2).pack(pady=20)

    def load_original(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.df_original = pd.read_excel(file_path)

    def load_transformed(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.df_transformed = pd.read_excel(file_path)

    def page2(self):
        if self.df_original is None or self.df_transformed is None:
            return

        for widget in self.root.winfo_children():
            widget.destroy()

        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        self.string_columns = [col for col in self.df_original.columns if self.df_original[col].dtype == object]
        self.numeric_columns = [col for col in self.df_original.columns if pd.api.types.is_numeric_dtype(self.df_original[col])]

        tk.Label(frame, text="Select String Columns:").pack()

        list_frame = tk.Frame(frame)
        list_frame.pack(fill="both", expand=True)

        scrollbar_y = tk.Scrollbar(list_frame, orient="vertical")
        scrollbar_x = tk.Scrollbar(list_frame, orient="horizontal")
        self.listbox = tk.Listbox(list_frame, selectmode="multiple", yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.config(command=self.listbox.yview)
        scrollbar_x.config(command=self.listbox.xview)

        for col in self.string_columns:
            self.listbox.insert(tk.END, col)

        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        tk.Button(frame, text="Submit", command=self.generate_table_page).pack(pady=10)

    def generate_table_page(self):
        selected_indices = self.listbox.curselection()
        self.selected_string_columns = [self.string_columns[i] for i in selected_indices]
        self.remaining_columns = [col for col in self.string_columns if col not in self.selected_string_columns]

        df1 = self.df_original.copy()
        df2 = self.df_transformed.copy()

        key_col = "_filter_key"
        df1[key_col] = df1[self.selected_string_columns].astype(str).agg(' | '.join, axis=1)
        df2[key_col] = df2[self.selected_string_columns].astype(str).agg(' | '.join, axis=1)

        merged = pd.merge(df1, df2, on=key_col, suffixes=('_original', '_transformed'))
        for num_col in self.numeric_columns:
            merged['status_' + num_col] = merged[f"{num_col}_original"].eq(merged[f"{num_col}_transformed"]).map({True: 'OK', False: 'Anomaly'})

        self.show_table(merged)

    def show_table(self, df):
        for widget in self.root.winfo_children():
            widget.destroy()

        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        filter_values = df['_filter_key'].unique().tolist()
        self.filter_var = tk.StringVar(value=filter_values)

        filter_frame = tk.Frame(frame)
        filter_frame.pack(fill="x")
        tk.Label(filter_frame, text="Filter:").pack(side="left")
        filter_listbox = tk.Listbox(filter_frame, listvariable=self.filter_var, selectmode="multiple", exportselection=False)
        filter_listbox.pack(side="left", fill="both", expand=True)

        scrollbar_y = tk.Scrollbar(filter_frame, orient="vertical", command=filter_listbox.yview)
        scrollbar_y.pack(side="left", fill="y")
        filter_listbox.config(yscrollcommand=scrollbar_y.set)

        tree_frame = tk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)

        scrollbar_y_tree = tk.Scrollbar(tree_frame, orient="vertical")
        scrollbar_x_tree = tk.Scrollbar(tree_frame, orient="horizontal")
        tree = ttk.Treeview(tree_frame, yscrollcommand=scrollbar_y_tree.set, xscrollcommand=scrollbar_x_tree.set)
        scrollbar_y_tree.config(command=tree.yview)
        scrollbar_x_tree.config(command=tree.xview)

        tree["columns"] = list(df.columns)
        tree["show"] = "headings"
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=150)

        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y_tree.grid(row=0, column=1, sticky="ns")
        scrollbar_x_tree.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        tk.Button(frame, text="Export to Excel", command=lambda: self.export_to_excel(df)).pack(pady=5)
        tk.Button(frame, text="Next Drill Down", command=lambda: self.page2_drill(df, filter_listbox)).pack(pady=5)

    def page2_drill(self, df, filter_listbox):
        selected_indices = filter_listbox.curselection()
        selected_values = [filter_listbox.get(i) for i in selected_indices]
        self.selected_values = selected_values if selected_values else None

        if self.selected_values:
            df = df[df['_filter_key'].isin(self.selected_values)]

        self.string_columns = self.remaining_columns
        self.page2()

    def export_to_excel(self, df):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not file_path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Reconciliation Table"

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Charts for each selected string column
        for col in self.selected_string_columns:
            if col in df.columns:
                ws_chart = wb.create_sheet(title=f"Chart_{col}")
                pivot = df.groupby(col)[[c for c in df.columns if "_original" in c or "_transformed" in c]].sum().reset_index()

                for r in dataframe_to_rows(pivot, index=False, header=True):
                    ws_chart.append(r)

                chart = BarChart()
                data = Reference(ws_chart, min_col=2, max_col=len(pivot.columns), min_row=1, max_row=len(pivot) + 1)
                cats = Reference(ws_chart, min_col=1, min_row=2, max_row=len(pivot) + 1)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(cats)
                chart.title = f"Original vs Transformed by {col}"
                ws_chart.add_chart(chart, "E5")

        wb.save(file_path)


root = tk.Tk()
app = ReconciliationApp(root)
root.mainloop()
```
