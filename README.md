```
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

class DataGroupingApp(ttk.Window):
    def __init__(self):
        super().__init__(title="Data Grouping Tool", themename="cosmo")
        self.geometry("950x600")
        self.datafile = None
        self.reqfile = None
        self.df = None

        self.page1()

    # ---------- PAGE 1 ----------
    def page1(self):
        for widget in self.winfo_children():
            widget.destroy()

        frame = ttk.Frame(self, padding=30)
        frame.pack(expand=True, fill="both")

        ttk.Label(frame, text="Select Input Files", font=("Helvetica", 18, "bold")).pack(pady=20)

        self.data_path_var = ttk.StringVar()
        self.req_path_var = ttk.StringVar()

        ttk.Label(frame, text="Select Data File (.xlsb):", bootstyle="primary").pack(anchor="w", pady=(10,0))
        ttk.Entry(frame, textvariable=self.data_path_var, width=60, state="readonly").pack(pady=5)
        ttk.Button(frame, text="Browse", bootstyle="info-outline", command=self.load_datafile).pack(pady=5)

        ttk.Label(frame, text="Select Requirement File (.xlsx):", bootstyle="primary").pack(anchor="w", pady=(20,0))
        ttk.Entry(frame, textvariable=self.req_path_var, width=60, state="readonly").pack(pady=5)
        ttk.Button(frame, text="Browse", bootstyle="info-outline", command=self.load_reqfile).pack(pady=5)

        self.next_btn = ttk.Button(frame, text="Next →", bootstyle="success", command=self.page2, state="disabled")
        self.next_btn.pack(pady=40)

    def load_datafile(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Binary Workbook", "*.xlsb")])
        if path:
            try:
                import pyxlsb
                self.df = pd.read_excel(path, engine="pyxlsb")
                self.data_path_var.set(path)
                self.datafile = path
                self.check_ready()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read xlsb file:\n{e}")

    def load_reqfile(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Workbook", "*.xlsx")])
        if path:
            self.req_path_var.set(path)
            self.reqfile = path
            self.check_ready()

    def check_ready(self):
        if self.datafile and self.reqfile:
            self.next_btn.config(state="normal")

    # ---------- PAGE 2 ----------
    def page2(self):
        for widget in self.winfo_children():
            widget.destroy()

        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Create New Column", font=("Helvetica", 18, "bold")).pack(pady=15)

        top_frame = ttk.Frame(frame)
        top_frame.pack(pady=10, fill="x")

        ttk.Label(top_frame, text="Select Column:", width=15).grid(row=0, column=0, sticky="w", padx=5)
        self.col_var = ttk.StringVar()
        col_menu = ttk.Combobox(top_frame, textvariable=self.col_var, values=list(self.df.columns), state="readonly")
        col_menu.grid(row=0, column=1, padx=5)
        col_menu.bind("<<ComboboxSelected>>", self.load_unique_values)

        ttk.Label(top_frame, text="New Column Name:", width=18).grid(row=0, column=2, sticky="w", padx=5)
        self.newcol_var = ttk.StringVar()
        ttk.Entry(top_frame, textvariable=self.newcol_var, width=25).grid(row=0, column=3, padx=5)

        # Unique values section
        mid_frame = ttk.Labelframe(frame, text="Unique Values", bootstyle="info", padding=10)
        mid_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.value_listbox = ttk.Listbox(mid_frame, selectmode="extended", height=15)
        self.value_listbox.pack(side="left", fill="both", expand=True, padx=10)

        scrollbar = ttk.Scrollbar(mid_frame, command=self.value_listbox.yview)
        scrollbar.pack(side="left", fill="y")
        self.value_listbox.config(yscrollcommand=scrollbar.set)

        right_frame = ttk.Frame(mid_frame)
        right_frame.pack(side="left", padx=10)

        ttk.Label(right_frame, text="Group Name:").pack(pady=5)
        self.group_name_var = ttk.StringVar()
        ttk.Entry(right_frame, textvariable=self.group_name_var).pack(pady=5)

        ttk.Button(right_frame, text="Add Group", bootstyle="primary", command=self.add_group).pack(pady=10)
        ttk.Button(right_frame, text="Create Column", bootstyle="success", command=self.create_column).pack(pady=10)
        ttk.Button(right_frame, text="Next →", bootstyle="info-outline", command=self.page2).pack(pady=10)

        self.groups = {}
        self.group_preview = ttk.Treeview(frame, columns=("Group", "Values"), show="headings", height=6)
        self.group_preview.heading("Group", text="Group Name")
        self.group_preview.heading("Values", text="Values")
        self.group_preview.pack(fill="x", padx=20, pady=10)

    def load_unique_values(self, event=None):
        self.value_listbox.delete(0, "end")
        col = self.col_var.get()
        if col:
            uniques = sorted(self.df[col].dropna().unique())
            for val in uniques:
                self.value_listbox.insert("end", str(val))
        self.groups = {}
        for item in self.group_preview.get_children():
            self.group_preview.delete(item)

    def add_group(self):
        group_name = self.group_name_var.get().strip()
        if not group_name:
            messagebox.showwarning("Missing Input", "Please enter a group name.")
            return
        selected = [self.value_listbox.get(i) for i in self.value_listbox.curselection()]
        if not selected:
            messagebox.showwarning("Missing Selection", "Select values to group.")
            return
        self.groups[group_name] = selected
        for i in reversed(self.value_listbox.curselection()):
            self.value_listbox.delete(i)
        self.group_preview.insert("", "end", values=(group_name, ", ".join(selected)))
        self.group_name_var.set("")

    def create_column(self):
        newcol = self.newcol_var.get().strip()
        col = self.col_var.get()
        if not newcol or not col or not self.groups:
            messagebox.showwarning("Incomplete", "Fill all fields and create groups first.")
            return

        mapping = {}
        for group, vals in self.groups.items():
            for v in vals:
                mapping[v] = group

        self.df[newcol] = self.df[col].map(mapping)
        messagebox.showinfo("Success", f"Column '{newcol}' created successfully!")

        # Refresh Page
        self.page2()

# Run the app
if __name__ == "__main__":
    app = DataGroupingApp()
    app.mainloop()
