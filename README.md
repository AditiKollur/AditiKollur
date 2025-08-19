```
import tkinter as tk
from tkinter import messagebox
import pandas as pd

class ColumnSelectorApp:
    def __init__(self, root, df):
        self.root = root
        self.df = df
        self.consol = None  # Will store selected columns
        self.root.title("Select Columns")
        self.root.geometry("400x400")

        tk.Label(root, text="Select Columns to Keep:", font=("Arial", 12)).pack(pady=10)

        # Listbox with multiple selection
        self.listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=40, height=15)
        for col in df.columns:
            self.listbox.insert(tk.END, col)
        self.listbox.pack(pady=10)

        # Confirm button
        tk.Button(root, text="Confirm Selection", bg="lightgreen", command=self.confirm_selection).pack(pady=10)

    def confirm_selection(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "Please select at least one column.")
            return

        selected_cols = [self.listbox.get(i) for i in selected_indices]
        self.consol = self.df[selected_cols].copy()

        # Close GUI
        self.root.destroy()

# Example usage
if __name__ == "__main__":
    df = pd.DataFrame({
        "sarid": [101, 102, 103],
        "box": ["BoxA", "BoxB", "BoxC"],
        "Jan25_YTD": [100, 200, 300],
        "Feb25_YTD": [150, 250, 350],
        "segment": ["Retail", "Corporate", "SME"]
    })

    root = tk.Tk()
    app = ColumnSelectorApp(root, df)
    root.mainloop()

    # consol contains the selected columns
    consol = app.consol
    print("Selected columns dataframe:")
    print(consol)


```
