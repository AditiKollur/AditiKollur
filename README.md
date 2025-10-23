```
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button, Label, Combobox, Frame
import pandas as pd


class DataSelectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Selection Portal")
        self.style = Style(theme="cosmo")

        self.data_file = None
        self.req_file = None
        self.df = None

        self.frame_main = Frame(root, padding=30)
        self.frame_main.pack(fill="both", expand=True)

        self.build_file_upload_ui()

    def clear_frame(self):
        for widget in self.frame_main.winfo_children():
            widget.destroy()

    def build_file_upload_ui(self):
        """UI for selecting Data and Req files"""
        self.clear_frame()

        Label(self.frame_main, text="📂 Upload Required Files", font=("Segoe UI", 16, "bold")).pack(pady=10)

        Button(self.frame_main, text="Select Data File", bootstyle="info", command=self.load_data_file, width=25).pack(pady=10)
        Button(self.frame_main, text="Select Req File", bootstyle="info", command=self.load_req_file, width=25).pack(pady=10)

        self.lbl_status = Label(self.frame_main, text="", font=("Segoe UI", 10))
        self.lbl_status.pack(pady=10)

        Button(self.frame_main, text="Submit Files", bootstyle="success", command=self.submit_files, width=20).pack(pady=20)

    def load_data_file(self):
        file_path = filedialog.askopenfilename(title="Select Data File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.data_file = file_path
            self.lbl_status.config(text="Data File Loaded ✅")
            try:
                self.df = pd.read_excel(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Error loading Excel file: {e}")

    def load_req_file(self):
        file_path = filedialog.askopenfilename(title="Select Req File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.req_file = file_path
            self.lbl_status.config(text="Req File Loaded ✅")

    def submit_files(self):
        if not self.data_file or not self.req_file:
            messagebox.showwarning("Missing Files", "Please upload both Data and Req files before proceeding.")
            return
        if self.df is None:
            messagebox.showerror("Error", "Data file not loaded properly.")
            return
        self.build_column_selection_ui()

    def build_column_selection_ui(self):
        """UI for selecting multiple pairs (and one triple) of columns"""
        self.clear_frame()

        Label(self.frame_main, text="📊 Column Mapping", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, columnspan=5, pady=20)

        rows = [
            ("Segment Performance", "segment", 2),
            ("Product Performance", "product", 2),
            ("Region Performance", "region", 3),
        ]

        self.vars = {}
        columns = list(self.df.columns)

        for i, (label_text, key, num_combos) in enumerate(rows, start=1):
            Label(self.frame_main, text=label_text, font=("Segoe UI", 12)).grid(row=i, column=0, padx=10, pady=10, sticky="w")

            combo_vars = []
            combo_boxes = []

            for j in range(num_combos):
                var = tk.StringVar()
                combo = Combobox(self.frame_main, textvariable=var, values=columns, bootstyle="info", width=20)

                # Disable all except the first
                if j > 0:
                    combo.config(state="disabled")

                combo.grid(row=i, column=j + 1, padx=10, pady=5)
                combo_vars.append(var)
                combo_boxes.append(combo)

            # Bind logic for progressive enablement and intra-row exclusion
            for idx, combo in enumerate(combo_boxes):
                combo.bind(
                    "<<ComboboxSelected>>",
                    lambda e, i=idx, cboxes=combo_boxes: self.handle_selection_in_row(i, cboxes),
                )

            self.vars[key] = combo_vars

        Button(self.frame_main, text="Submit Selection", bootstyle="success", command=self.submit_selection, width=20).grid(
            row=len(rows) + 1, column=0, columnspan=5, pady=30
        )

    def handle_selection_in_row(self, index, combo_boxes):
        """Mutual exclusion + progressive unlock"""
        all_columns = list(self.df.columns)
        selected = [cb.get() for cb in combo_boxes if cb.get()]

        # Update mutual exclusion
        for cb in combo_boxes:
            current_val = cb.get()
            new_values = [c for c in all_columns if c not in selected or c == current_val]
            cb.config(values=new_values)

        # Unlock the next dropdown only if current is selected
        if index + 1 < len(combo_boxes):
            next_cb = combo_boxes[index + 1]
            if combo_boxes[index].get():
                next_cb.config(state="normal")
            else:
                # If user clears previous dropdown, lock again
                next_cb.config(state="disabled")
                next_cb.set("")  # reset value

    def submit_selection(self):
        """Collect all selections and validate"""
        selected_data = {}

        for key, var_list in self.vars.items():
            selected_values = [v.get() for v in var_list]
            if any(not v for v in selected_values):
                messagebox.showwarning("Incomplete Selection", f"Please complete selection for {key.capitalize()} Performance.")
                return
            selected_data[key] = selected_values

        # Summary message
        summary_lines = []
        for k, vals in selected_data.items():
            summary_lines.append(f"{k.capitalize()}: " + " | ".join(vals))

        messagebox.showinfo("Selection Saved ✅", "\n".join(summary_lines))
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = DataSelectionApp(root)
    root.geometry("850x500")
    root.mainloop()
