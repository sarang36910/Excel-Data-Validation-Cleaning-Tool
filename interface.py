import tkinter as tk
from tkinter import filedialog, messagebox
import os

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)

def run_validation():
    input_file = entry_file.get()
    if not os.path.isfile(input_file):
        messagebox.showerror("Error", "Invalid input file path")
        return
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx")],
                                               initialfile="validated_output.xlsx")
    if not output_file:
        return
    # Call your validation function here
    try:
        run_validation_all(input_file, output_file)
        messagebox.showinfo("Success", f"Validation completed and saved to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Validation failed: {str(e)}")

app = tk.Tk()
app.title("Excel Validator")

tk.Label(app, text="Select Excel File:").grid(row=0, column=0, padx=10, pady=10)
entry_file = tk.Entry(app, width=50)
entry_file.grid(row=0, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=select_file).grid(row=0, column=2, padx=10, pady=10)

tk.Button(app, text="Run Validation", command=run_validation).grid(row=1, column=1, pady=20)

app.mainloop()
