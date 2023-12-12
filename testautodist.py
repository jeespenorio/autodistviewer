import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

# Function to import Excel file and display data
def import_and_display():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            display_data(df)
        except Exception as e:
            print("Error:", e)

# Function to display data in a tkinter window
def display_data(df):
    # Create a tkinter window
    root = tk.Tk()
    root.title("Excel Data Display")

    # Function to export data to a CSV file
    def export_data():
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if filepath:
            df.to_csv(filepath, index=False)

    # Create a treeview widget
    tree = ttk.Treeview(root)

    # Define columns
    tree["columns"] = tuple(df.columns)
    tree.heading("#0", text="Index")
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, stretch=tk.YES)

    # Insert data into the treeview
    for i, row in df.iterrows():
        tree.insert("", "end", text=i, values=tuple(row))

    tree.pack(expand=tk.YES, fill=tk.BOTH)

    # Add Export button
    export_button = ttk.Button(root, text="Export", command=export_data)
    export_button.pack()

    root.mainloop()

# Create a tkinter window for importing
root_import = tk.Tk()
root_import.title("Import Excel File")

# Add Import button
import_button = ttk.Button(root_import, text="Import Excel File", command=import_and_display)
import_button.pack()

root_import.mainloop()
