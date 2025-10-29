```
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ttkbootstrap as tb
import pandas as pd


class DataMappingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Mapping Application")
        self.style = tb.Style("flatly")
        self.root.geometry("950x700")

        # Variables
        self.data_file = None
        self.req_file = None
        self.df_data = None
        self.df_req = None
        self.new_column_details = {}
        self.selections = {}
        self.groups = {}
        self.remaining_values = []

        # UI setup
        self.main_frame = ttk.Frame(self.root, padding=20)
        self.main_frame.pack(fill="both", expand=True)
        self.show_step1()

    # ---------------- STEP 1: FILE UPLOAD ---------------- #
    def show_step1(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        ttk.Label(self.main_frame, text="Step 1: Upload Files", font=("Segoe UI", 14, "bold")).pack(pady=10)

        frame = ttk.Frame(self.main_frame)
        frame.pack(pady=20)

        ttk.Button(frame, text="Upload Data File", command=self.load_data_file, bootstyle="primary").grid(row=0, column=0, padx=10)
        ttk.Button(frame, text="Upload Requirement File", command=self.load_req_file, bootstyle="info").grid(row=0, column=1, padx=10)

        ttk.Button(self.main_frame, text="Next →", bootstyle="success", command=self.show_step2).pack(pady=20)

    def load_data_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")])
        if path:
            self.data_file = path
            if path.endswith(".csv"):
                self.df_data = pd.read_csv(path)
            else:
                self.df_data = pd.read_excel(path)
            messagebox.showinfo("Success", f"Data file loaded: {path.split('/')[-1]}")

    def load_req_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")])
        if path:
            self.req_file = path
            if path.endswith(".csv"):
                self.df_req = pd.read_csv(path)
            else:
                self.df_req = pd.read_excel(path)
            messagebox.showinfo("Success", f"Requirement file loaded: {path.split('/')[-1]}")

    # ---------------- STEP 2: COLUMN CREATION ---------------- #
    def show_step2(self):
        if self.df_data is None or self.df_req is None:
            messagebox.showwarning("Error", "Please upload both files first!")
            return

        for widget in self.main_frame.winfo_children():
            widget.destroy()

        ttk.Label(self.main_frame, text="Step 2: Build Custom Columns", font=("Segoe UI", 14, "bold")).pack(pady=10)

        ttk.Label(self.main_frame, text="Select base column:").pack(pady=5)
        self.base_col_var = tk.StringVar()
        self.base_col_cb = ttk.Combobox(self.main_frame, textvariable=self.base_col_var, values=list(self.df_data.columns))
        self.base_col_cb.pack(pady=5)

        ttk.Button(self.main_frame, text="Load Unique Values", bootstyle="secondary", command=self.load_unique_values).pack(pady=10)

        self.groups_frame = ttk.Frame(self.main_frame)
        self.groups_frame.pack(fill="x", pady=10)

        ttk.Button(self.main_frame, text="Next →", bootstyle="success", command=self.show_step3).pack(pady=20)

    def load_unique_values(self):
        base_col = self.base_col_var.get()
        if not base_col:
            messagebox.showwarning("Error", "Please select a base column.")
            return

        unique_values = sorted(self.df_data[base_col].dropna().unique().tolist())
        self.remaining_values = unique_values.copy()

        for widget in self.groups_frame.winfo_children():
            widget.destroy()

        ttk.Label(self.groups_frame, text=f"Unique values from '{base_col}':").pack()
        ttk.Label(self.groups_frame, text=str(unique_values), wraplength=900).pack(pady=5)

        ttk.Label(self.groups_frame, text="Enter new column name:").pack(pady=5)
        self.new_col_name = tk.StringVar()
        ttk.Entry(self.groups_frame, textvariable=self.new_col_name).pack(pady=5)

        ttk.Button(self.groups_frame, text="Add Group", bootstyle="info", command=self.add_group).pack(pady=10)
        ttk.Button(self.groups_frame, text="Replicate Group Pattern", bootstyle="warning", command=self.replicate_groups).pack(pady=10)
        ttk.Button(self.groups_frame, text="Create Column", bootstyle="success", command=self.create_new_column).pack(pady=10)

        self.group_widgets = []

    def add_group(self):
        frame = ttk.Labelframe(self.groups_frame, text=f"Group {len(self.group_widgets)+1}")
        frame.pack(fill="x", pady=5)

        group_name = tk.StringVar()
        ttk.Label(frame, text="Group Name:").grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(frame, textvariable=group_name).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Select values (Ctrl+Click for multiple):").grid(row=1, column=0, columnspan=2, padx=5, pady=5)
        listbox = tk.Listbox(frame, selectmode="multiple", height=6, exportselection=False)
        for v in [v for v in self.remaining_values if v not in sum(self.groups.values(), [])]:
            listbox.insert(tk.END, v)
        listbox.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        ttk.Button(frame, text="Add to Group", bootstyle="success",
                   command=lambda: self.save_group(group_name, listbox)).grid(row=3, column=0, columnspan=2, pady=5)
        self.group_widgets.append((group_name, listbox))

    def save_group(self, name_var, listbox):
        name = name_var.get().strip()
        selected_indices = listbox.curselection()
        selected_values = [listbox.get(i) for i in selected_indices]

        if not name or not selected_values:
            messagebox.showwarning("Error", "Please specify a group name and select at least one value.")
            return

        if name not in self.groups:
            self.groups[name] = []
        self.groups[name].extend(selected_values)

        # Remove used values
        for val in selected_values:
            if val in self.remaining_values:
                self.remaining_values.remove(val)

        messagebox.showinfo("Added", f"Added {len(selected_values)} values to group '{name}'.")

        # Refresh all remaining listboxes
        for _, lb in self.group_widgets:
            lb.delete(0, tk.END)
            for v in [v for v in self.remaining_values if v not in sum(self.groups.values(), [])]:
                lb.insert(tk.END, v)

    def replicate_groups(self):
        if not self.groups:
            messagebox.showwarning("Error", "No groups created to replicate.")
            return
        for name in list(self.groups.keys()):
            new_name = f"{name}_copy"
            if self.remaining_values:
                self.groups[new_name] = [self.remaining_values.pop(0)]
        messagebox.showinfo("Done", "Groups replicated for remaining items!")

    def create_new_column(self):
        base_col = self.base_col_var.get()
        new_col = self.new_col_name.get()
        if not base_col or not new_col or not self.groups:
            messagebox.showwarning("Error", "Please complete all details before creating column.")
            return

        mapping = self.groups
        self.df_data[new_col] = self.df_data[base_col].apply(lambda x: next((k for k, v in mapping.items() if x in v), x))
        self.new_column_details = {"base_column": base_col, "new_column_name": new_col, "mapping": mapping}
        messagebox.showinfo("Success", f"Column '{new_col}' created successfully!")

    # ---------------- STEP 3: COLUMN SELECTION ---------------- #
    def show_step3(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        ttk.Label(self.main_frame, text="Step 3: Column Selection", font=("Segoe UI", 14, "bold")).pack(pady=10)
        frame = ttk.Frame(self.main_frame)
        frame.pack(pady=10)

        mapping_levels = {
            "Segment Performance": 2,
            "Product Performance": 2,
            "Region Performance": 3
        }

        for r, (level, count) in enumerate(mapping_levels.items()):
            ttk.Label(frame, text=level).grid(row=r, column=0, padx=5, pady=5)
            options = list(self.df_data.columns)
            row_sel = {}
            for c in range(1, count + 1):
                var = tk.StringVar()
                cb = ttk.Combobox(frame, values=options, textvariable=var)
                cb.grid(row=r, column=c, padx=5, pady=5)
                row_sel[f"Column {c}"] = var
            self.selections[level] = row_sel

        ttk.Button(self.main_frame, text="Finish", bootstyle="success", command=self.finish_app).pack(pady=20)

    def finish_app(self):
        summary = f"Data File: {self.data_file}\nRequirement File: {self.req_file}\n\n"
        summary += f"New Column: {self.new_column_details}\nSelections: {self.selections}"
        messagebox.showinfo("Summary", summary)
        self.root.destroy()


if __name__ == "__main__":
    root = tb.Window(themename="flatly")
    app = DataMappingApp(root)
    root.mainloop()
