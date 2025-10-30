```
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog, StringVar, Listbox, END, MULTIPLE
import pandas as pd


class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("Custom Column Builder")
        self.geometry("850x600")
        self.resizable(False, False)

        # Internal state
        self.data_file = None
        self.req_file = None
        self.df = None
        self.group_mapping = {}

        # Notebook setup
        self.notebook = ttk.Notebook(self)
        self.page1 = ttk.Frame(self.notebook)
        self.page2 = ttk.Frame(self.notebook)
        self.notebook.add(self.page1, text="Step 1 – Select Files")
        self.notebook.add(self.page2, text="Step 2 – Create Custom Column")
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Build UI pages
        self.build_page1()
        self.build_page2()

        # Disable Page 2 initially
        self.notebook.tab(1, state="disabled")

    # ---------------- PAGE 1 -----------------
    def build_page1(self):
        frame = ttk.Labelframe(self.page1, text="1️⃣ File Selection", padding=20)
        frame.pack(fill=X, padx=30, pady=50)

        # Data file
        ttk.Label(frame, text="Select Data File (.csv/.xlsx/.xls/.xlsb):").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Button(frame, text="Browse", bootstyle=PRIMARY, command=self.select_data_file).grid(row=0, column=1, padx=10)
        self.data_label = ttk.Label(frame, text="No data file selected", width=60)
        self.data_label.grid(row=0, column=2, padx=10, pady=5)

        # Requirement file
        ttk.Label(frame, text="Select Requirement File (.xlsx):").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Button(frame, text="Browse", bootstyle=INFO, command=self.select_req_file).grid(row=1, column=1, padx=10)
        self.req_label = ttk.Label(frame, text="No requirement file selected", width=60)
        self.req_label.grid(row=1, column=2, padx=10, pady=5)

        # Load button
        ttk.Button(frame, text="Load Data", bootstyle=SUCCESS, command=self.load_data).grid(
            row=2, column=0, columnspan=3, pady=30
        )

    # ---------------- PAGE 2 -----------------
    def build_page2(self):
        frame = ttk.Labelframe(self.page2, text="2️⃣ Create Custom Column", padding=20)
        frame.pack(fill=BOTH, expand=True, padx=20, pady=20)

        # Variables
        self.col_var = StringVar()
        self.group_name_var = StringVar()

        # Dropdown for columns
        ttk.Label(frame, text="Select Column:").grid(row=0, column=0, sticky=W)
        self.col_dropdown = ttk.Combobox(frame, textvariable=self.col_var, state="readonly", width=30)
        self.col_dropdown.grid(row=0, column=1, padx=10, pady=5)

        # Group name entry
        ttk.Label(frame, text="Enter Group Name:").grid(row=1, column=0, sticky=W)
        ttk.Entry(frame, textvariable=self.group_name_var, width=30).grid(row=1, column=1, padx=10, pady=5)

        # Value listbox
        self.value_listbox = Listbox(frame, selectmode=MULTIPLE, width=40, height=10)
        self.value_listbox.grid(row=2, column=0, columnspan=2, pady=10)

        # Buttons
        ttk.Button(frame, text="Load Values", bootstyle=INFO, command=self.load_values).grid(row=3, column=0, pady=10)
        ttk.Button(frame, text="Add Group", bootstyle=SUCCESS, command=self.add_group).grid(row=3, column=1, pady=10)
        ttk.Button(frame, text="Replicate Remaining", bootstyle=WARNING, command=self.replicate_groups).grid(
            row=4, column=0, columnspan=2, pady=10
        )

    # ---------------- FILE HANDLERS -----------------
    def select_data_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[("Excel/CSV Files", "*.csv *.xlsx *.xls *.xlsb")]
        )
        if file_path:
            self.data_file = file_path
            self.data_label.config(text=file_path.split("/")[-1])

    def select_req_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Requirement File",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if file_path:
            self.req_file = file_path
            self.req_label.config(text=file_path.split("/")[-1])

    def load_data(self):
        if not self.data_file or not self.req_file:
            Messagebox.show_error("Please select both data and requirement files.", "Missing Files")
            return

        try:
            # Load Data File
            if self.data_file.endswith(".xlsb"):
                self.df = pd.read_excel(self.data_file, engine="pyxlsb")
            elif self.data_file.endswith((".xlsx", ".xls")):
                self.df = pd.read_excel(self.data_file)
            else:
                self.df = pd.read_csv(self.data_file)

            # Populate dropdown
            self.col_dropdown["values"] = list(self.df.columns)

            # Unlock Page 2
            self.notebook.tab(1, state="normal")
            self.notebook.select(1)
            Messagebox.show_info("Files loaded successfully. You can now proceed to Step 2.", "Success")

        except Exception as e:
            Messagebox.show_error(f"Error loading file:\n{e}", "Load Error")

    # ---------------- PAGE 2 FUNCTIONS -----------------
    def load_values(self):
        col = self.col_var.get()
        if not col:
            Messagebox.show_warning("Please select a column first.", "No Column Selected")
            return
        values = sorted(self.df[col].dropna().unique().tolist())
        self.value_listbox.delete(0, END)
        for v in values:
            self.value_listbox.insert(END, v)

    def add_group(self):
        group_name = self.group_name_var.get().strip()
        selected = [self.value_listbox.get(i) for i in self.value_listbox.curselection()]

        if not group_name:
            Messagebox.show_warning("Enter a group name.", "Missing Name")
            return
        if not selected:
            Messagebox.show_warning("Select at least one value.", "No Selection")
            return

        self.group_mapping[group_name] = selected
        Messagebox.show_info(f"Group '{group_name}' added with {len(selected)} items.", "Group Added")

        for i in reversed(self.value_listbox.curselection()):
            self.value_listbox.delete(i)

    def replicate_groups(self):
        remaining = self.value_listbox.get(0, END)
        for v in remaining:
            self.group_mapping[v] = [v]
        self.value_listbox.delete(0, END)
        Messagebox.show_info("Remaining values replicated as individual groups.", "Replication Complete")


if __name__ == "__main__":
    app = App()
    app.mainloop()
