```
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class DataSelectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Column Builder & Selector")
        self.style = ttk.Style("cosmo")

        # state
        self.data_file = None
        self.req_file = None
        self.df = None
        self.groups = {}
        self.created_columns = []
        self.column_selection = {}

        self.build_file_selection_ui()

    # ----------------------- STEP 1: FILE SELECTION -------------------------
    def build_file_selection_ui(self):
        self.clear_window()
        frame = ttk.Frame(self.root, padding=25)
        frame.pack(fill=BOTH, expand=True)

        ttk.Label(frame, text="Step 1: Select Data Files", font=("Helvetica", 14, "bold")).pack(pady=10)

        ttk.Button(frame, text="Select Data File", command=self.load_data_file, bootstyle=PRIMARY).pack(pady=5)
        ttk.Button(frame, text="Select Req File", command=self.load_req_file, bootstyle=INFO).pack(pady=5)

        self.file_label = ttk.Label(frame, text="", bootstyle="secondary")
        self.file_label.pack(pady=10)

        ttk.Button(frame, text="Next → Build Custom Columns", command=self.goto_column_creation, bootstyle=SUCCESS).pack(pady=20)

    def load_data_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.data_file = path
            self.df = pd.read_excel(path)
            self.file_label.config(text=f"Data File Loaded: {path.split('/')[-1]}")

    def load_req_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.req_file = path
            self.file_label.config(text=f"Req File Loaded: {path.split('/')[-1]}")

    # ----------------------- STEP 2: COLUMN CREATION -------------------------
    def goto_column_creation(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please load a Data File first.")
            return

        self.clear_window()
        frame = ttk.Frame(self.root, padding=25)
        frame.pack(fill=BOTH, expand=True)

        ttk.Label(frame, text="Step 2: Create Custom Columns", font=("Helvetica", 14, "bold")).pack(pady=10)

        self.col_names = list(self.df.columns)
        self.used_values = set()
        self.groups = {}

        # --- Base column + new col name ---
        top_frame = ttk.Frame(frame)
        top_frame.pack(fill=X, pady=10)

        ttk.Label(top_frame, text="Base Column:", width=14).pack(side=LEFT)
        self.base_col_var = tk.StringVar()
        self.base_col_cb = ttk.Combobox(top_frame, textvariable=self.base_col_var, values=self.col_names, state="readonly", width=25)
        self.base_col_cb.pack(side=LEFT, padx=5)
        self.base_col_cb.bind("<<ComboboxSelected>>", self.load_unique_values)

        ttk.Label(top_frame, text="New Column Name:").pack(side=LEFT, padx=5)
        self.new_col_name_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.new_col_name_var, width=25).pack(side=LEFT, padx=5)

        # --- Unique values listbox ---
        ttk.Label(frame, text="Select values for each group:", font=("Helvetica", 11, "bold")).pack(pady=(10, 5))
        self.unique_listbox = tk.Listbox(frame, selectmode="multiple", height=8, exportselection=False)
        self.unique_listbox.pack(fill=X, pady=5)

        # --- Group name + add button ---
        group_frame = ttk.Frame(frame)
        group_frame.pack(fill=X, pady=5)

        ttk.Label(group_frame, text="Group Name:", width=14).pack(side=LEFT)
        self.group_name_var = tk.StringVar()
        ttk.Entry(group_frame, textvariable=self.group_name_var, width=25).pack(side=LEFT, padx=5)
        ttk.Button(group_frame, text="Add Group", command=self.add_group, bootstyle=PRIMARY).pack(side=LEFT, padx=5)

        # --- Buttons ---
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="Create Column", command=self.create_custom_column, bootstyle=SUCCESS).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Next → Column Selection", command=self.goto_column_selection, bootstyle=INFO).pack(side=LEFT, padx=5)

    def load_unique_values(self, event=None):
        col = self.base_col_var.get()
        if not col:
            return
        uniques = [v for v in self.df[col].dropna().unique() if v not in self.used_values]
        self.unique_listbox.delete(0, tk.END)
        for u in uniques:
            self.unique_listbox.insert(tk.END, u)

    def add_group(self):
        group_name = self.group_name_var.get().strip()
        if not group_name:
            messagebox.showwarning("Warning", "Enter a group name.")
            return

        selected = [self.unique_listbox.get(i) for i in self.unique_listbox.curselection()]
        if not selected:
            messagebox.showwarning("Warning", "Select at least one value for this group.")
            return

        self.groups[group_name] = selected
        self.used_values.update(selected)
        self.load_unique_values()

        messagebox.showinfo("Group Added", f"Group '{group_name}' created with {len(selected)} items.")
        self.group_name_var.set("")

    def create_custom_column(self):
        base_col = self.base_col_var.get()
        new_col = self.new_col_name_var.get().strip()

        if not base_col or not new_col or not self.groups:
            messagebox.showwarning("Warning", "Please fill all fields and add at least one group.")
            return

        mapping = {}
        for gname, vals in self.groups.items():
            for v in vals:
                mapping[v] = gname

        self.df[new_col] = self.df[base_col].map(mapping).fillna(self.df[base_col])
        self.col_names.append(new_col)
        self.created_columns.append(new_col)

        messagebox.showinfo("Success", f"Column '{new_col}' created successfully!")

        # Reset for next possible creation
        self.used_values.clear()
        self.groups.clear()
        self.load_unique_values()

    # ----------------------- STEP 3: COLUMN SELECTION -------------------------
    def goto_column_selection(self):
        self.clear_window()
        frame = ttk.Frame(self.root, padding=25)
        frame.pack(fill=BOTH, expand=True)

        ttk.Label(frame, text="Step 3: Select Columns for Each Level", font=("Helvetica", 14, "bold")).pack(pady=10)

        self.selection_vars = {}
        rows = [
            ("Segment Performance", 2),
            ("Product Performance", 2),
            ("Region Performance", 3),
        ]

        for row_label, num_cols in rows:
            row_frame = ttk.Frame(frame)
            row_frame.pack(fill=X, pady=8)
            ttk.Label(row_frame, text=row_label, width=22).pack(side=LEFT)

            row_vars = []
            for i in range(num_cols):
                var = tk.StringVar()
                cb = ttk.Combobox(row_frame, textvariable=var, values=self.col_names, width=25, state="readonly")
                cb.pack(side=LEFT, padx=5)
                row_vars.append(var)

            # Prevent same-column selection within a row
            for var in row_vars:
                var.trace_add("write", lambda *a, v=row_vars: self.update_row_options(v))

            self.selection_vars[row_label] = row_vars

        ttk.Button(frame, text="Finish", command=self.finish, bootstyle=SUCCESS).pack(pady=20)

    def update_row_options(self, row_vars):
        selected = [v.get() for v in row_vars if v.get()]
        for v in row_vars:
            current = v.get()
            v.widget.config(values=[c for c in self.col_names if c not in selected or c == current])

    # ----------------------- STEP 4: FINISH -------------------------
    def finish(self):
        self.column_selection = {
            row: [v.get() for v in vars if v.get()] for row, vars in self.selection_vars.items()
        }

        messagebox.showinfo("Done", "Selections saved successfully.")
        print("✅ Data File:", self.data_file)
        print("✅ Req File:", self.req_file)
        print("✅ Created Columns:", self.created_columns)
        print("✅ Column Selections:", self.column_selection)
        print("✅ Groups Used:", self.groups)
        print("✅ Final Columns in DataFrame:", list(self.df.columns))

        self.root.destroy()

    # ----------------------- UTIL -------------------------
    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()

# ----------------------- RUN APP -------------------------
if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = DataSelectionApp(root)
    root.mainloop()
