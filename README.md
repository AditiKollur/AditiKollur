```
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class DataSelectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data & Column Selection Tool")
        self.style = ttk.Style("cosmo")

        # state variables
        self.data_file = None
        self.req_file = None
        self.df = None
        self.column_selection = {}
        self.groups = {}
        self.used_values = set()

        self.build_file_selection_ui()

    # ----------------------- STEP 1: FILE SELECTION -------------------------
    def build_file_selection_ui(self):
        self.clear_window()
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill=BOTH, expand=True)

        ttk.Label(frame, text="Step 1: Select Data Files", font=("Helvetica", 14, "bold")).pack(pady=10)

        ttk.Button(frame, text="Select Data File", command=self.load_data_file, bootstyle=PRIMARY).pack(pady=5)
        ttk.Button(frame, text="Select Req File", command=self.load_req_file, bootstyle=INFO).pack(pady=5)

        self.file_label = ttk.Label(frame, text="", bootstyle="secondary")
        self.file_label.pack(pady=10)

        ttk.Button(frame, text="Next → Column Selection", command=self.goto_column_selection, bootstyle=SUCCESS).pack(pady=20)

    def load_data_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.data_file = file_path
            self.df = pd.read_excel(file_path)
            self.file_label.config(text=f"Data File Loaded: {file_path.split('/')[-1]}")

    def load_req_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.req_file = file_path
            self.file_label.config(text=f"Req File Loaded: {file_path.split('/')[-1]}")

    # ----------------------- STEP 2: COLUMN SELECTION & COLUMN CREATION -------------------------
    def goto_column_selection(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please load Data File first.")
            return

        self.clear_window()
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill=BOTH, expand=True)

        ttk.Label(frame, text="Step 2: Select Columns & Create Custom Columns", font=("Helvetica", 14, "bold")).pack(pady=10)

        self.col_names = list(self.df.columns)
        self.selection_vars = {}

        rows = [
            ("Segment Performance", 2),
            ("Product Performance", 2),
            ("Region Performance", 3),
        ]

        # --- Column Selection Table ---
        table_frame = ttk.Frame(frame)
        table_frame.pack(pady=15, fill=X)

        for row_label, num_cols in rows:
            row_frame = ttk.Frame(table_frame)
            row_frame.pack(fill=X, pady=8)
            ttk.Label(row_frame, text=row_label, width=22).pack(side=LEFT)

            row_vars = []
            for i in range(num_cols):
                var = tk.StringVar()
                cb = ttk.Combobox(row_frame, textvariable=var, values=self.col_names, width=25, state="readonly")
                cb.pack(side=LEFT, padx=5)
                row_vars.append(var)

            for i, var in enumerate(row_vars):
                var.trace_add("write", lambda *args, vars=row_vars: self._update_row_options(vars))

            self.selection_vars[row_label] = row_vars

        # --- Divider ---
        ttk.Separator(frame, bootstyle="secondary").pack(fill=X, pady=15)

        # --- Column Creation UI ---
        ttk.Label(frame, text="➕ Create a Custom Column", font=("Helvetica", 12, "bold")).pack(pady=5)

        creation_frame = ttk.Frame(frame)
        creation_frame.pack(fill=X, pady=10)

        ttk.Label(creation_frame, text="Base Column:", width=14).pack(side=LEFT)
        self.base_col_var = tk.StringVar()
        self.base_col_cb = ttk.Combobox(creation_frame, textvariable=self.base_col_var, values=self.col_names, state="readonly", width=20)
        self.base_col_cb.pack(side=LEFT, padx=5)
        self.base_col_cb.bind("<<ComboboxSelected>>", self.load_unique_values)

        ttk.Label(creation_frame, text="New Column Name:").pack(side=LEFT, padx=5)
        self.new_col_name_var = tk.StringVar()
        ttk.Entry(creation_frame, textvariable=self.new_col_name_var, width=20).pack(side=LEFT, padx=5)

        ttk.Button(creation_frame, text="Add Group", command=self.add_group, bootstyle=PRIMARY).pack(side=LEFT, padx=5)

        ttk.Label(frame, text="Select Unique Values for Group:").pack()
        self.unique_listbox = tk.Listbox(frame, selectmode="multiple", height=7, exportselection=False)
        self.unique_listbox.pack(fill=X, pady=8)

        ttk.Button(frame, text="Create Column", command=self.create_custom_column, bootstyle=SUCCESS).pack(pady=5)

        # --- Finish Button ---
        ttk.Button(frame, text="Finish Selection", command=self.finish, bootstyle=SUCCESS).pack(pady=20)

    def _update_row_options(self, vars):
        selected = [v.get() for v in vars if v.get()]
        for v in vars:
            current_val = v.get()
            v.widget.config(values=[c for c in self.col_names if c not in selected or c == current_val])

    # --- Column Creation Handlers ---
    def load_unique_values(self, event=None):
        col = self.base_col_var.get()
        if not col:
            return
        uniques = [v for v in self.df[col].dropna().unique() if v not in self.used_values]
        self.unique_listbox.delete(0, tk.END)
        for u in uniques:
            self.unique_listbox.insert(tk.END, u)

    def add_group(self):
        selected = [self.unique_listbox.get(i) for i in self.unique_listbox.curselection()]
        if not selected:
            messagebox.showwarning("Warning", "Select at least one value for the group.")
            return
        group_name = self.new_col_name_var.get()
        if not group_name:
            messagebox.showwarning("Warning", "Enter a new column name first.")
            return

        if group_name not in self.groups:
            self.groups[group_name] = []

        self.groups[group_name].append(selected)
        self.used_values.update(selected)
        self.load_unique_values()

        messagebox.showinfo("Group Added", f"Added {len(selected)} items to group '{group_name}'.")

    def create_custom_column(self):
        col = self.base_col_var.get()
        group_name = self.new_col_name_var.get()

        if not col or group_name not in self.groups:
            messagebox.showwarning("Warning", "Incomplete column creation.")
            return

        # Build mapping dictionary
        mapping = {}
        for idx, group_values in enumerate(self.groups[group_name]):
            for v in group_values:
                mapping[v] = f"{group_name}_Group{idx+1}"

        # Create new column
        self.df[group_name] = self.df[col].map(mapping).fillna(self.df[col])
        self.col_names.append(group_name)

        messagebox.showinfo("Success", f"Custom column '{group_name}' added successfully!")

        # Refresh all dropdowns
        for row_vars in self.selection_vars.values():
            for var in row_vars:
                var.widget.config(values=self.col_names)

    # ----------------------- STEP 3: FINISH -------------------------
    def finish(self):
        self.column_selection = {
            row: [v.get() for v in vars if v.get()] for row, vars in self.selection_vars.items()
        }
        messagebox.showinfo("Done", "Selections saved successfully.")
        print("Data File:", self.data_file)
        print("Req File:", self.req_file)
        print("Column Selections:", self.column_selection)
        print("Groups:", self.groups)
        print("Final Columns:", list(self.df.columns))
        self.root.destroy()

    # ----------------------- UTILITIES -------------------------
    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = DataSelectionApp(root)
    root.mainloop()
