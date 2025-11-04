```
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import StringVar, Listbox, END, MULTIPLE, filedialog, messagebox
import pandas as pd


dict_func = {
    "Function_A": 2,
    "Function_B": 3,
    "Function_C": 4
}


class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("Custom Column Builder")
        self.geometry("1100x700")  # Increased window size
        self.resizable(False, False)

        # internal state
        self.data_file = None
        self.req_file = None
        self.df = None
        self.req_df = None
        self.group_mapping = {}
        self.mappings_df = None

        # notebook setup
        self.notebook = ttk.Notebook(self)
        self.page1 = ttk.Frame(self.notebook)
        self.page2 = ttk.Frame(self.notebook)
        self.page3 = ttk.Frame(self.notebook)

        self.notebook.add(self.page1, text="Step 1 – Select Files")
        self.notebook.add(self.page2, text="Step 2 – Create Custom Column")
        self.notebook.add(self.page3, text="Step 3 – Function Mapping")
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Disable tabs 2 & 3 initially
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
        ttk.Button(frame, text="Browse", bootstyle=PRIMARY, command=self.select_data_file).grid(row=0, column=1, padx=10)
        self.data_label = ttk.Label(frame, text="No data file selected", width=60)
        self.data_label.grid(row=0, column=2, padx=10, pady=5)

        ttk.Label(frame, text="Select Requirement File (.xlsx):").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Button(frame, text="Browse", bootstyle=INFO, command=self.select_req_file).grid(row=1, column=1, padx=10)
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

        ttk.Button(frame, text="Next → Function Mapping", bootstyle=INFO,
                   command=self.proceed_to_page3).grid(row=6, column=0, columnspan=4, pady=10)

    # ---------------- PAGE 3 -----------------
    def build_page3(self):
        self.mapping_frame = ttk.Labelframe(self.page3, text="3️⃣ Function Mapping", padding=20)
        self.mapping_frame.pack(fill=BOTH, expand=True, padx=20, pady=20)

        headers = ["Column", "Functionality", "Level 1", "Level 2", "Level 3", "Level 4"]
        for i, h in enumerate(headers):
            ttk.Label(self.mapping_frame, text=h).grid(row=0, column=i, padx=10, pady=5)

        self.rows = []
        for r in range(4):
            col_entry = ttk.Entry(self.mapping_frame, width=15)
            col_entry.insert(0, f"Header {r + 1}")
            col_entry.grid(row=r + 1, column=0, padx=5, pady=5)

            func_var = StringVar()
            func_combo = ttk.Combobox(self.mapping_frame, textvariable=func_var,
                                      values=list(dict_func.keys()), width=15, state="readonly")
            func_combo.grid(row=r + 1, column=1, padx=5, pady=5)

            level_vars, level_combos = [], []
            for j in range(4):
                lvl_var = StringVar()
                lvl_combo = ttk.Combobox(self.mapping_frame, textvariable=lvl_var,
                                         width=15, state="disabled")
                lvl_combo.grid(row=r + 1, column=j + 2, padx=5, pady=5)
                level_vars.append(lvl_var)
                level_combos.append(lvl_combo)

            func_combo.bind("<<ComboboxSelected>>",
                            lambda e, lvls=level_combos, fvar=func_var: self.populate_levels(lvls, fvar))

            self.rows.append((col_entry, func_var, func_combo, level_vars, level_combos))

        self.submit_btn = ttk.Button(self.mapping_frame, text="Submit", bootstyle=SUCCESS,
                                     command=self.submit_mapping, state="disabled")
        self.submit_btn.grid(row=6, column=0, columnspan=6, pady=20)

    # ---------------- LOGIC -----------------
    def select_data_file(self):
        path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[("All Supported", "*.csv *.xlsx *.xls *.xlsb")]
        )
        if path:
            self.data_file = path
            self.data_label.config(text=path)

    def select_req_file(self):
        path = filedialog.askopenfilename(title="Select Requirement File", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.req_file = path
            self.req_label.config(text=path)

    def load_data(self):
        if not self.data_file or not self.req_file:
            messagebox.showerror("Missing File", "Please select both files.")
            return
        try:
            ext = self.data_file.split(".")[-1]
            if ext == "csv":
                self.df = pd.read_csv(self.data_file)
            elif ext in ["xlsx", "xls"]:
                self.df = pd.read_excel(self.data_file)
            elif ext == "xlsb":
                self.df = pd.read_excel(self.data_file, engine="pyxlsb")
            else:
                raise ValueError("Unsupported format")
            self.req_df = pd.read_excel(self.req_file)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        string_cols = self.df.select_dtypes(include="object").columns.tolist()
        self.col_combo.config(values=string_cols)
        self.notebook.tab(1, state="normal")
        messagebox.showinfo("Success", "Files loaded successfully!")

    def load_unique_values(self):
        col = self.col_var.get()
        if not col:
            messagebox.showerror("Error", "Select a column.")
            return
        self.listbox.delete(0, END)
        vals = self.df[col].dropna().unique().tolist()
        for v in vals:
            self.listbox.insert(END, v)
        self.group_mapping.clear()

    def load_selected_group(self):
        sel = [self.listbox.get(i) for i in self.listbox.curselection()]
        group = self.group_entry.get().strip()
        if not sel or not group:
            messagebox.showerror("Error", "Select values and enter a group name.")
            return
        for v in sel:
            self.group_mapping[v] = group
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)

    def replicate_remaining(self):
        for v in self.listbox.get(0, END):
            self.group_mapping[v] = v
        self.listbox.delete(0, END)

    def create_new_column(self):
        if not self.group_mapping:
            messagebox.showerror("Error", "Define at least one group.")
            return
        col = self.col_var.get()
        new_col = self.new_name.get().strip() or f"{col}_group"
        self.df[new_col] = self.df[col].map(self.group_mapping).fillna(self.df[col])
        messagebox.showinfo("Success", f"New column '{new_col}' created!")

    def proceed_to_page3(self):
        if self.df is None:
            messagebox.showerror("Error", "Load data first.")
            return
        self.notebook.tab(2, state="normal")
        self.notebook.select(2)

    def populate_levels(self, level_combos, func_var):
        for combo in level_combos:
            combo.set("")
            combo.config(state="disabled")

        func = func_var.get()
        n_levels = dict_func.get(func, 0)
        available_cols = self.df.select_dtypes(include="object").columns.tolist()

        for i in range(n_levels):
            combo = level_combos[i]
            combo.config(values=available_cols, state="readonly")
            combo.set(f"Level {i + 1}")
            combo.bind("<<ComboboxSelected>>",
                       lambda e, idx=i, lvls=level_combos: self.update_next_levels(idx, lvls, available_cols))

        self.validate_submit_button()

    def update_next_levels(self, idx, level_combos, available_cols):
        selected_vals = [combo.get() for combo in level_combos if combo.get()]
        for i in range(idx + 1, len(level_combos)):
            if level_combos[i].cget("state") != "disabled":
                remaining = [c for c in available_cols if c not in selected_vals]
                level_combos[i].config(values=remaining)
        self.validate_submit_button()

    def validate_submit_button(self):
        enable = False
        for row in self.rows:
            func = row[1].get()
            if func:
                n_levels = dict_func.get(func, 0)
                if all(row[3][i].get() for i in range(n_levels)):
                    enable = True
                else:
                    enable = False
                    break
        self.submit_btn.config(state="normal" if enable else "disabled")

    def submit_mapping(self):
        data = []
        for row in self.rows:
            col_name = row[0].get()
            func = row[1].get()
            if not func:
                continue
            n_levels = dict_func.get(func, 0)
            levels = [row[3][i].get() for i in range(n_levels)]
            if all(levels):
                d = {"Column": col_name, "Functionality": func}
                for i, lvl in enumerate(levels):
                    d[f"Level{i + 1}"] = lvl
                data.append(d)

        if not data:
            messagebox.showerror("Error", "No valid selections made.")
            return

        self.mappings_df = pd.DataFrame(data)
        self.destroy()  # close window


def launch_gui():
    app = App()
    app.mainloop()
    return getattr(app, "df", None), getattr(app, "mappings_df", None)
