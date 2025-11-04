```
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import StringVar, Listbox, END, MULTIPLE, filedialog, messagebox
import pandas as pd


class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("Custom Column Builder")
        self.geometry("900x700")
        self.resizable(False, False)

        # internal state
        self.data_file = None
        self.req_file = None
        self.df = None
        self.req_df = None
        self.group_mapping = {}

        # dict_func setup
        self.dict_func = {
            "Function_A": 2,
            "Function_B": 3,
            "Function_C": 4
        }

        # notebook setup
        self.notebook = ttk.Notebook(self)
        self.page1 = ttk.Frame(self.notebook)
        self.page2 = ttk.Frame(self.notebook)
        self.page3 = ttk.Frame(self.notebook)

        self.notebook.add(self.page1, text="Step 1 – Select Files")
        self.notebook.add(self.page2, text="Step 2 – Create Custom Column")
        self.notebook.add(self.page3, text="Step 3 – Advanced Mapping")
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Disable 2nd and 3rd tab initially
        self.notebook.tab(1, state="disabled")
        self.notebook.tab(2, state="disabled")

        self.build_page1()
        self.build_page2()
        self.build_page3()

    # ---------------- PAGE 1 -----------------
    def build_page1(self):
        frame = ttk.Labelframe(self.page1, text="1️⃣ File Selection", padding=20)
        frame.pack(fill=X, padx=30, pady=50)

        ttk.Label(frame, text="Select Data File (.csv/.xlsx/.xls/.xlsb):").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Button(frame, text="Browse", bootstyle=PRIMARY,
                   command=self.select_data_file).grid(row=0, column=1, padx=10)
        self.data_label = ttk.Label(frame, text="No data file selected", width=60)
        self.data_label.grid(row=0, column=2, padx=10, pady=5)

        ttk.Label(frame, text="Select Requirement File (.xlsx):").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Button(frame, text="Browse", bootstyle=INFO,
                   command=self.select_req_file).grid(row=1, column=1, padx=10)
        self.req_label = ttk.Label(frame, text="No requirement file selected", width=60)
        self.req_label.grid(row=1, column=2, padx=10, pady=5)

        ttk.Button(frame, text="Load Data", bootstyle=SUCCESS, command=self.load_data).grid(
            row=2, column=0, columnspan=3, pady=30
        )

    # ---------------- PAGE 2 -----------------
    def build_page2(self):
        frame = ttk.Labelframe(self.page2, text="2️⃣ Create Custom Column", padding=20)
        frame.pack(fill=BOTH, expand=True, padx=20, pady=20)

        ttk.Label(frame, text="Select Column:").grid(row=0, column=0, sticky=W, pady=5)
        self.col_var = StringVar()
        self.col_combo = ttk.Combobox(frame, textvariable=self.col_var, width=25, state="readonly")
        self.col_combo.grid(row=0, column=1, padx=10)

        ttk.Label(frame, text="New Name:").grid(row=0, column=2, sticky=W, padx=20)
        self.new_name = ttk.Entry(frame, width=20)
        self.new_name.grid(row=0, column=3)

        ttk.Button(frame, text="Load Unique Values", bootstyle=INFO,
                   command=self.load_unique_values).grid(row=1, column=0, columnspan=4, pady=10)

        self.listbox = Listbox(frame, height=10, width=40, selectmode=MULTIPLE)
        self.listbox.grid(row=2, column=0, columnspan=2, rowspan=3, padx=10, pady=10)

        ttk.Label(frame, text="Group Name:").grid(row=2, column=2, sticky=W)
        self.group_entry = ttk.Entry(frame, width=20)
        self.group_entry.grid(row=2, column=3, pady=5)

        ttk.Button(frame, text="Load", bootstyle=SECONDARY,
                   command=self.load_selected_group).grid(row=3, column=3, pady=5)
        ttk.Button(frame, text="Replicate", bootstyle=WARNING,
                   command=self.replicate_remaining).grid(row=3, column=2, pady=5)

        ttk.Button(frame, text="Create New Column", bootstyle=SUCCESS,
                   command=self.create_new_column).grid(row=5, column=0, columnspan=4, pady=20)

        ttk.Button(frame, text="Proceed →", bootstyle=PRIMARY,
                   command=self.proceed_to_page3).grid(row=6, column=0, columnspan=4, pady=10)

    # ---------------- PAGE 3 -----------------
    def build_page3(self):
        self.page3_frame = ttk.Labelframe(self.page3, text="3️⃣ Advanced Mapping", padding=20)
        self.page3_frame.pack(fill=BOTH, expand=True, padx=20, pady=20)

        ttk.Label(self.page3_frame, text="Configure additional mappings (4 rows)").grid(row=0, column=0, columnspan=5, pady=10)

        self.mapping_entries = []

        for i in range(4):
            row_dict = {}
            # Column text entry
            col_var = StringVar()
            col_entry = ttk.Entry(self.page3_frame, textvariable=col_var, width=20)
            col_entry.grid(row=i + 1, column=0, padx=5, pady=5)
            row_dict["column_name"] = col_var

            # Functionality dropdown
            func_var = StringVar()
            func_combo = ttk.Combobox(self.page3_frame, textvariable=func_var,
                                      values=list(self.dict_func.keys()), state="readonly", width=20)
            func_combo.grid(row=i + 1, column=1, padx=5)
            func_combo.bind("<<ComboboxSelected>>",
                            lambda e, row=i: self.load_levels_for_function(row))
            row_dict["functionality"] = func_var
            row_dict["levels"] = []

            self.mapping_entries.append(row_dict)

        self.submit_btn = ttk.Button(self.page3_frame, text="Submit", bootstyle=SUCCESS,
                                     command=self.submit_mappings, state="disabled")
        self.submit_btn.grid(row=6, column=0, columnspan=5, pady=20)

    def load_levels_for_function(self, row):
        """Dynamically add level dropdowns based on selected function."""
        row_info = self.mapping_entries[row]
        func_name = row_info["functionality"].get()

        # remove old level dropdowns if exist
        for level in row_info["levels"]:
            level["widget"].destroy()
        row_info["levels"].clear()

        if not func_name:
            return

        num_levels = self.dict_func[func_name]
        all_options = list(self.df.columns)

        def on_select(current_level):
            """Unlock next level and filter values."""
            selected_values = [lvl["var"].get() for lvl in row_info["levels"] if lvl["var"].get()]
            if current_level + 1 < len(row_info["levels"]):
                next_combo = row_info["levels"][current_level + 1]["widget"]
                next_combo.config(state="readonly",
                                  values=[v for v in all_options if v not in selected_values])

            # Enable submit if at least one selection made
            if any(lvl["var"].get() for lvl in row_info["levels"]):
                self.submit_btn.config(state="normal")

        for j in range(num_levels):
            lvl_var = StringVar()
            lvl_combo = ttk.Combobox(self.page3_frame, textvariable=lvl_var, state="disabled", width=15)
            lvl_combo.grid(row=row + 1, column=2 + j, padx=5)
            lvl_combo.set(f"Level{j + 1}")
            lvl_combo.bind("<<ComboboxSelected>>", lambda e, lvl=j: on_select(lvl))

            row_info["levels"].append({"var": lvl_var, "widget": lvl_combo})

        # Unlock first level
        row_info["levels"][0]["widget"].config(state="readonly", values=all_options)

    def submit_mappings(self):
        """Collect mappings and close GUI"""
        mappings = []
        for row in self.mapping_entries:
            col_name = row["column_name"].get().strip()
            func = row["functionality"].get()
            levels = [lvl["var"].get() for lvl in row["levels"] if lvl["var"].get()]
            if col_name and func and levels:
                mappings.append({
                    "Column": col_name,
                    "Functionality": func,
                    "Levels": levels
                })

        if not mappings:
            messagebox.showerror("Error", "Please make selections in at least one row.")
            return

        self.mappings_df = pd.DataFrame(mappings)
        print("\n✅ Final Mapping DataFrame:")
        print(self.mappings_df)
        messagebox.showinfo("Done", "Mappings submitted successfully!")
        self.destroy()

    # ---------------- LOGIC -----------------
    def select_data_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[("All Supported", "*.csv *.xlsx *.xls *.xlsb"),
                       ("CSV files", "*.csv"),
                       ("Excel files", "*.xlsx *.xls *.xlsb")]
        )
        if file_path:
            self.data_file = file_path
            self.data_label.config(text=file_path)

    def select_req_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Requirement File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.req_file = file_path
            self.req_label.config(text=file_path)

    def load_data(self):
        if not self.data_file or not self.req_file:
            messagebox.showerror("Missing File", "Please select both files before loading.")
            return

        try:
            ext = self.data_file.split(".")[-1].lower()
            if ext == "csv":
                self.df = pd.read_csv(self.data_file)
            elif ext in ["xlsx", "xls"]:
                self.df = pd.read_excel(self.data_file)
            elif ext == "xlsb":
                self.df = pd.read_excel(self.data_file, engine="pyxlsb")
        except Exception as e:
            messagebox.showerror("Error Loading Data File", str(e))
            return

        try:
            self.req_df = pd.read_excel(self.req_file)
        except Exception as e:
            messagebox.showerror("Error Loading Requirement File", str(e))
            return

        self.col_combo.config(values=list(self.df.columns))
        self.notebook.tab(1, state="normal")
        messagebox.showinfo("Success", f"✅ Loaded data ({len(self.df)} rows, {len(self.df.columns)} columns).")

    def load_unique_values(self):
        if self.df is None:
            messagebox.showerror("Error", "Load data first.")
            return
        col = self.col_var.get()
        if not col:
            messagebox.showerror("Error", "Select a column first.")
            return
        self.listbox.delete(0, END)
        unique_vals = self.df[col].dropna().unique().tolist()
        for val in unique_vals:
            self.listbox.insert(END, str(val))
        self.group_mapping.clear()

    def load_selected_group(self):
        selected_indices = self.listbox.curselection()
        selected_values = [self.listbox.get(i) for i in selected_indices]
        group_name = self.group_entry.get().strip()
        if not selected_values or not group_name:
            messagebox.showerror("Missing Input", "Select values and enter a group name.")
            return

        for val in selected_values:
            self.group_mapping[val] = group_name
        for i in reversed(selected_indices):
            self.listbox.delete(i)

    def replicate_remaining(self):
        remaining = self.listbox.get(0, END)
        for val in remaining:
            self.group_mapping[val] = val
        self.listbox.delete(0, END)

    def create_new_column(self):
        if self.df is None or not self.group_mapping:
            messagebox.showerror("Error", "No data or groups defined.")
            return
        col = self.col_var.get()
        new_col = self.new_name.get().strip() or f"{col}_group"
        self.df[new_col] = self.df[col].map(self.group_mapping).fillna(self.df[col])
        messagebox.showinfo("Done", f"New column '{new_col}' added successfully!")

    def proceed_to_page3(self):
        if self.df is None:
            messagebox.showerror("Error", "Please load data first.")
            return
        self.notebook.tab(2, state="normal")
        self.notebook.select(2)


if __name__ == "__main__":
    app = App()
    app.mainloop()
