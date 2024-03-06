import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

class ExcelMerger:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel Sheet Merger")
        self.master.geometry("700x600")

        # Description
        ttk.Label(master, text="This application merges selected sheets from multiple Excel files.",
                  font=("Helvetica", 16)).grid(row=0, column=0, columnspan=2, padx=20, pady=(20, 10))

        # "Select Data" Button
        ttk.Label(master, text="Step 1: Select Excel Files").grid(row=1, column=0, sticky=tk.W, padx=20)
        self.select_button = ttk.Button(master, text="Select Data", command=self.load_files)
        self.select_button.grid(row=1, column=1, sticky=tk.W, padx=20, pady=(5, 20))

        # Listbox for Sheet Names
        ttk.Label(master, text="Step 2: Select Sheets to Merge (Choose one option)").grid(row=2, column=0, sticky=tk.W, padx=20)
        ttk.Label(master, text="Option 1: Select from list").grid(row=3, column=0, sticky=tk.W, padx=20)
        self.sheet_var = tk.StringVar(value=[" "])
        self.sheet_listbox = tk.Listbox(master, listvariable=self.sheet_var, selectmode=tk.MULTIPLE, height=10, width=30)
        self.sheet_listbox.grid(row=4, column=0, padx=20, pady=(5, 0))

        # Dropdown for Sheet Names
        ttk.Label(master, text="Option 2: Select from dropdown").grid(row=3, column=1, sticky=tk.W, padx=20)
        self.sheet_combo = ttk.Combobox(master, values=[" "], height=10, width=30)
        self.sheet_combo.grid(row=4, column=1, padx=20, pady=(5, 0))

        # Text Entry for Sheet Names
        ttk.Label(master, text="Option 3: Enter sheet names separated by commas").grid(row=5, column=0, columnspan=2, sticky=tk.W, padx=20)
        self.sheet_entry = ttk.Entry(master, width=30)
        self.sheet_entry.grid(row=6, column=0, columnspan=2, padx=20, pady=(10, 20))

        # Select All / Deselect All Buttons
        self.select_all_button = ttk.Button(master, text="Select All", command=self.select_all_sheets)
        self.select_all_button.grid(row=7, column=0, padx=10, pady=(0, 10))
        self.deselect_all_button = ttk.Button(master, text="Deselect All", command=self.deselect_all_sheets)
        self.deselect_all_button.grid(row=7, column=1, padx=10, pady=(0, 10))

        # Merge Button
        ttk.Label(master, text="Step 3: Merge Sheets").grid(row=8, column=0, sticky=tk.W, padx=20)
        self.merge_button = ttk.Button(master, text="Merge Sheets", command=self.merge_sheets)
        self.merge_button.grid(row=8, column=1, sticky=tk.W, padx=20, pady=(5, 20))

        # Internal variables
        self.files = []
        self.sheets = {}

    def load_files(self):
        self.files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if not self.files:
            return
        self.sheets = {}
        for file in self.files:
            xls = pd.ExcelFile(file)
            self.sheets[file] = xls.sheet_names

        all_sheet_names = set()
        for sheet_list in self.sheets.values():
            all_sheet_names.update(sheet_list)

        sorted_sheet_names = list(sorted(all_sheet_names))
        self.sheet_var.set([" "] + sorted_sheet_names)
        self.sheet_combo['values'] = [" "] + sorted_sheet_names

    def select_all_sheets(self):
        self.sheet_listbox.selection_set(0, tk.END)

    def deselect_all_sheets(self):
        self.sheet_listbox.selection_clear(0, tk.END)

    def merge_sheets(self):
        selected_sheets = list({self.sheet_listbox.get(i) for i in self.sheet_listbox.curselection()})
        selected_sheets += self.sheet_combo.get().split(",")
        selected_sheets += self.sheet_entry.get().split(",")
        selected_sheets = list(set(selected_sheets))
        selected_sheets = [x for x in selected_sheets if x.strip()]  # Remove empty or blank selections

        if not selected_sheets:
            return

        merged_data = {}
        for sheet in selected_sheets:
            merged_data[sheet] = []
            for file in self.files:
                if sheet in self.sheets[file]:
                    df = pd.read_excel(file, sheet_name=sheet)
                    merged_data[sheet].append(df)

        for sheet, data_list in merged_data.items():
            merged_data[sheet] = pd.concat(data_list, ignore_index=True) if data_list else None

        # Save file
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            return

        with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
            for sheet, df in merged_data.items():
                if df is not None:
                    df.to_excel(writer, sheet_name=sheet, index=False)

        messagebox.showinfo("Success", "Sheets merged successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMerger(root)
    root.mainloop()
