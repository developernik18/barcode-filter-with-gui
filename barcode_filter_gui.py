import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# --- FUNCTIONS ---
def filter_barcodes():
    try:
        # Get file paths from user inputs
        input_file = input_entry.get()
        barcode_file = barcode_entry.get()
        output_file = output_entry.get()
        barcode_column_main_name = main_column_entry.get()
        barcode_column_list_name = list_column_entry.get()

        # Read Excel files
        df_main = pd.read_excel(input_file)
        df_barcodes = pd.read_excel(barcode_file)

        # Convert barcode list to a Python list
        barcodes_to_keep = df_barcodes[barcode_column_list_name].astype(str).tolist()

        # Filter main DataFrame
        filtered_df = df_main[df_main[barcode_column_main_name].astype(str).isin(barcodes_to_keep)]

        # Save output
        filtered_df.to_excel(output_file, index=False)
        messagebox.showinfo("Success", f"Filtered rows saved to {output_file}\nTotal rows kept: {len(filtered_df)}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def browse_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def browse_save(entry):
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)

# --- GUI ---
root = tk.Tk()
root.title("Excel Barcode Filter")

# Labels and entries
tk.Label(root, text="Main Excel File:").grid(row=0, column=0, sticky="e")
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1)
tk.Button(root, text="Browse", command=lambda: browse_file(input_entry)).grid(row=0, column=2)

tk.Label(root, text="Barcode List File:").grid(row=1, column=0, sticky="e")
barcode_entry = tk.Entry(root, width=50)
barcode_entry.grid(row=1, column=1)
tk.Button(root, text="Browse", command=lambda: browse_file(barcode_entry)).grid(row=1, column=2)

tk.Label(root, text="Output File:").grid(row=2, column=0, sticky="e")
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=2, column=1)
tk.Button(root, text="Browse", command=lambda: browse_save(output_entry)).grid(row=2, column=2)

tk.Label(root, text="Main Barcode Column:").grid(row=3, column=0, sticky="e")
main_column_entry = tk.Entry(root, width=20)
main_column_entry.grid(row=3, column=1, sticky="w")
main_column_entry.insert(0, "barcode")  # default value

tk.Label(root, text="Barcode List Column:").grid(row=4, column=0, sticky="e")
list_column_entry = tk.Entry(root, width=20)
list_column_entry.grid(row=4, column=1, sticky="w")
list_column_entry.insert(0, "barcode")  # default value

# Filter button
tk.Button(root, text="Filter Barcodes", command=filter_barcodes, bg="green", fg="white").grid(row=5, column=0, columnspan=3, pady=10)

root.mainloop()
