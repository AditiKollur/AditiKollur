```
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *


class DataMappingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Data Mapping Tool")
        self.style = ttk.Style("flatly")

        # Stored data
        self.data_file = None
        self.req_file = None
        self.df_data = None
        self.df_req = None
        self.new_column_details = {}
        self.selections = {}

        self.build_ui()

    # ---------- MAIN UI ----------
    def build_ui(self):
        ttk.Label(self.root, text="📁 Step 1: Upload Files", font=("Segoe UI", 14, "bold")).pack(pady=10)

        frame_upload = ttk.Frame(self.root)
        frame_upload.pack(pady=5)

        ttk.Button(frame_upload, text="Select Data File", bootstyle=PRIMARY, command=self.load_data_file).grid(row=0, column=0, padx=5)
        ttk.Button(frame_upload, text="Select Requirement File", bootstyle=PRIMARY, command=self.load_req_file).grid(row=0, column=1, padx=5)

        self.label_files = ttk.Label(self.root, text="No files selected", font=("Segoe UI", 10))
        self.label_files.pack(pady=5)

        ttk.Separator(self.root, bootstyle="info").pack(fill="x", pady=10)

        # Step 2 - Optional Column Builder
        ttk.Label(self.root, text="🧩 Step 2: Optional - Build Custom Column", font=("Segoe UI", 14, "bold")).pack(pady=10)

        self.frame_build = ttk.Labelframe(self.root, text="Build Column from Existing", bootstyle="info")
        self.frame_build.pack(fill="x", padx=10, pady=5)
        self.build_column_ui()

        ttk.Separator(self.root, bootstyle="info").pack(fill="x", pady=10)

        # Step 3 - Column Mapping
        ttk.Label(self.root, text="🗂️ Step 3: Column Mappings", font=("Segoe UI", 14, "bold")).pack(pady=10)
        self.frame_mapping = ttk.Frame(self.root)
        self.frame_mapping.pack(fill="x", padx=10, pady=5)

        self.build_mapping_ui()

        ttk.Button(self.root, text="Submit All", bootstyle=SUCCESS, command=self.submit_all).pack(pady=15)

    # ---------- LOAD FILES ----------
    def load_data_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV or Excel files", "*.csv *.xlsx")])
        if file_path:
            self.data_file = file_path
            if file_path.endswith(".csv"):
                self.df_data = pd.read_csv(file_path)
            else:
                self.df_data = pd.read_excel(file_path)
            self.update_file_label()

    def load_req_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV or Excel files", "*.csv *.xlsx")])
        if file_path:
            self.req_file = file_path
            if file_path.endswith(".csv"):
                self.df_req = pd.read_csv(file_path)
            else:
                self.df_req = pd.read_excel(file_path)
            self.update_file_label()

    def update_file_label(self):
        if self.data_file and self.req_file:
            self.label_files.config(text=f"✅ Data: {self.data_file.split('/')[-1]} | Req: {self.req_file.split('/')[-1]}")
        elif self.data_file:
            self.label_files.config(text=f"Data File Loaded: {self.data_file.split('/')[-1]}")
        elif self.req_file:
            self.label_files.config(text=f"Requirement File Loaded: {self.req_file.split('/')[-1]}")

    # ---------- COLUMN BUILDER ----------
    def build_column_ui(self):
        ttk.Label(self.frame_build, text="Base Column:").grid(row=0, column=0, padx=5, pady=5)
        self.cmb_base_col = ttk.Combobox(self.frame_build, state="readonly")
        self.cmb_base_col.grid(row=0, column=1, padx=5, pady=5)

        ttk.Button(self.frame_build, text="Load Columns", bootstyle=INFO, command=self.load_columns_for_builder).grid(row=0, column=2, padx=5)

        ttk.Label(self.frame_build, text="New Column Name:").grid(row=1, column=0, padx=5, pady=5)
        self.entry_new_col = ttk.Entry(self.frame_build)
        self.entry_new_col.grid(row=1, column=1, padx=5, pady=5)

        self.frame_groups = ttk.Frame(self.frame_build)
        self.frame_groups.grid(row=2, column=0, columnspan=3, pady=5)
        self.group_entries = []
        self.add_group_row()

        ttk.Button(self.frame_build, text="Add Group", bootstyle=SECONDARY, command=self.add_group_row).grid(row=3, column=0, pady=5)
        ttk.Button(self.frame_build, text="Create Column", bootstyle=SUCCESS, command=self.create_new_column).grid(row=3, column=1, pady=5)

    def load_columns_for_builder(self):
        if self.df_data is not None:
            self.cmb_base_col["values"] = list(self.df_data.columns)
        else:
            messagebox.showerror("Error", "Please load a data file first.")

    def add_group_row(self):
        row = len(self.group_entries)
        frame_row = ttk.Frame(self.frame_groups)
        frame_row.pack(fill="x", pady=2)

        lbl = ttk.Label(frame_row, text=f"Group {row+1}:")
        lbl.pack(side="left", padx=5)

        cmb = ttk.Combobox(frame_row, width=25)
        cmb.pack(side="left", padx=5)
        name = ttk.Entry(frame_row, width=20)
        name.pack(side="left", padx=5)
        self.group_entries.append((cmb, name))

    def create_new_column(self):
        if self.df_data is None:
            messagebox.showerror("Error", "Please upload a data file first.")
            return

        base_col = self.cmb_base_col.get()
        new_col = self.entry_new_col.get()
        if not base_col or not new_col:
            messagebox.showerror("Error", "Please specify both base and new column names.")
            return

        unique_vals = self.df_data[base_col].dropna().unique().tolist()
        mapping = {}
        for cmb, name in self.group_entries:
            if cmb.get() and name.get():
                mapping[name.get()] = [cmb.get()]

        if not mapping:
            messagebox.showerror("Error", "Please define at least one group.")
            return

        self.df_data[new_col] = self.df_data[base_col].apply(
            lambda x: next((k for k, v in mapping.items() if x in v), x)
        )
        self.new_column_details = {"base_column": base_col, "new_column_name": new_col, "mapping": mapping}
        messagebox.showinfo("Success", f"✅ New column '{new_col}' created successfully!")

    # ---------- MAPPING SECTION ----------
    def build_mapping_ui(self):
        self.mapping_structure = {
            "Segment": ["Col 2", "Col 3"],
            "Product": ["Col 2", "Col 3"],
            "Region": ["Col 2", "Col 3", "Col 4"],
        }

        self.mapping_dropdowns = {}

        for idx, (level, cols) in enumerate(self.mapping_structure.items()):
            ttk.Label(self.frame_mapping, text=f"{level} Performance:", font=("Segoe UI", 11, "bold")).grid(row=idx, column=0, sticky="w", pady=5)
            row_widgets = []
            used_vars = set()

            for j, col_name in enumerate(cols, start=1):
                var = tk.StringVar()
                cmb = ttk.Combobox(self.frame_mapping, textvariable=var, state="readonly", width=18)
                cmb.grid(row=idx, column=j, padx=5, pady=3)

                def callback(event, this_var=var, row=row_widgets, used=used_vars):
                    val = this_var.get()
                    for c, _v in row:
                        if c != this_var:
                            c_values = [x for x in list(self.df_data.columns) if x != val]
                            c["values"] = c_values

                cmb.bind("<<ComboboxSelected>>", callback)
                row_widgets.append((var, cmb))
            self.mapping_dropdowns[level] = row_widgets

        ttk.Button(self.frame_mapping, text="Load Columns", bootstyle=INFO, command=self.load_columns_for_mapping).grid(row=len(self.mapping_structure)+1, column=1, pady=5)

    def load_columns_for_mapping(self):
        if self.df_data is None:
            messagebox.showerror("Error", "Please load a data file first.")
            return
        for level, widgets in self.mapping_dropdowns.items():
            for var, cmb in widgets:
                cmb["values"] = list(self.df_data.columns)

    # ---------- SUBMIT ----------
    def submit_all(self):
        if not self.data_file or not self.req_file:
            messagebox.showerror("Error", "Please upload both files first.")
            return

        self.selections = {}
        for level, widgets in self.mapping_dropdowns.items():
            self.selections[level] = [v.get() for v, _ in widgets if v.get()]

        summary = f"""
✅ Data File: {self.data_file}
✅ Requirement File: {self.req_file}
📄 New Column: {self.new_column_details.get('new_column_name', 'None')}
🗂️ Selections: {self.selections}
"""
        messagebox.showinfo("Summary", summary)


if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = DataMappingApp(root)
    root.mainloop()

app.data_file
app.req_file
app.new_column_details
app.selections
