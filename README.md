```
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import StringVar, Listbox, END, MULTIPLE, filedialog
import pandas as pd


class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("Custom Column Builder")
        self.geometry("820x600")
        self.resizable(False, False)

        # internal state
        self.data_file = None
        self.req_file = None
        self.df = None
        self.req_df = None
        self.group_mapping = {}

        # notebook setup
        self.notebook = ttk.Notebook(self)
        self.page1 = ttk.Frame(self.notebook)
        self.page2 = ttk.Frame(self.notebook)

        self.notebook.add(self.page1, text="Step 1 – Select Files")
        self.notebook.add(self.page2, text="Step 2 – Create Custom Column")
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Disable 2nd tab initially
        self.notebook.tab(1, state="disabled")

        self.build_page1()
        self.build_page2()

    # ---------------- PAGE 1 -----------------
    def build_page1(self):
        frame = ttk.Labelframe(self.page1, text="1️⃣ File Selection", padding=20)
        frame.pack(fill=X, padx=30, pady=50)

        # Data File
        ttk.Label(frame, text="Select Data File (.csv/.xlsx/.xls/.xlsb):").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Button(frame, text="Browse", bootstyle=PRIMARY,
                   command=self.select_data_file).grid(row=0, column=1, padx=10)
        self.data_label = ttk.Label(frame, text="No data file selected", width=60)
        self.data_label.grid(row=0, column=2, padx=10, pady=5)

        # Requirement File
        ttk.Label(frame, text="Select Requirement File (.xlsx):").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Button(frame, text="Browse", bootstyle=INFO,
                   command=self.select_req_file).grid(row=1, column=1, padx=10)
        self.req_label = ttk.Label(frame, text="No requirement file selected", width=60)
        self.req_label.grid(row=1, column=2, padx=10, pady=5)

        # Load Button
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

    # ---------------- LOGIC -----------------
    def select_data_file(self):
        """Pick main data file"""
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
        """Pick requirement Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Requirement File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.req_file = file_path
            self.req_label.config(text=file_path)

    def load_data(self):
        """Load both data and requirement files"""
        if not self.data_file or not self.req_file:
            ttk.Messagebox.show_error("Missing File", "Please select both files before loading.")
            return

        # --- Load main data file ---
        try:
            ext = self.data_file.split(".")[-1].lower()
            if ext == "csv":
                self.df = pd.read_csv(self.data_file)
            elif ext in ["xlsx", "xls"]:
                self.df = pd.read_excel(self.data_file)
            elif ext == "xlsb":
                self.df = pd.read_excel(self.data_file, engine="pyxlsb")
            else:
                raise ValueError("Unsupported file format.")
        except Exception as e:
            ttk.Messagebox.show_error("Error Loading Data File", str(e))
            return

        # --- Load requirement file ---
        try:
            self.req_df = pd.read_excel(self.req_file)
        except Exception as e:
            ttk.Messagebox.show_error("Error Loading Requirement File", str(e))
            return

        # --- Enable next page ---
        self.col_combo.config(values=list(self.df.columns))
        self.notebook.tab(1, state="normal")  # unlock 2nd page
        ttk.Messagebox.show_info("Success", f"Loaded {len(self.df)} rows and {len(self.df.columns)} columns.\n"
                                            f"Requirement file loaded successfully.")
        print("✅ Data & Requirement files loaded.")

    def load_unique_values(self):
        if self.df is None:
            ttk.Messagebox.show_error("Error", "Load data first.")
            return
        col = self.col_var.get()
        if not col:
            ttk.Messagebox.show_error("Error", "Select a column first.")
            return
        self.listbox.delete(0, END)
        unique_vals = self.df[col].dropna().unique().tolist()
        for val in unique_vals:
            self.listbox.insert(END, str(val))
        self.group_mapping.clear()
        print(f"🔹 Unique values from '{col}': {unique_vals}")

    def load_selected_group(self):
        selected_indices = self.listbox.curselection()
        selected_values = [self.listbox.get(i) for i in selected_indices]
        group_name = self.group_entry.get().strip()
        if not selected_values or not group_name:
            ttk.Messagebox.show_error("Missing Input", "Select values and enter a group name.")
            return

        for val in selected_values:
            self.group_mapping[val] = group_name
        for i in reversed(selected_indices):
            self.listbox.delete(i)
        print(f"✅ {selected_values} → {group_name}")

    def replicate_remaining(self):
        remaining = self.listbox.get(0, END)
        for val in remaining:
            self.group_mapping[val] = val
        self.listbox.delete(0, END)
        print(f"🌀 Replicated remaining as self-groups → {remaining}")

    def create_new_column(self):
        if self.df is None or not self.group_mapping:
            ttk.Messagebox.show_error("Error", "No data or groups defined.")
            return
        col = self.col_var.get()
        new_col = self.new_name.get().strip() or f"{col}_group"
        self.df[new_col] = self.df[col].map(self.group_mapping).fillna(self.df[col])
        print(f"\n📊 Created new column '{new_col}' with group mappings:")
        print(self.group_mapping)
        ttk.Messagebox.show_info("Done", f"New column '{new_col}' added successfully!")


if __name__ == "__main__":
    app = App()
    app.mainloop()
