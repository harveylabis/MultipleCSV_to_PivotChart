import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import create
import glob
import os

def browse_folder():
    folder_path = filedialog.askdirectory(title="Select Folder Containing CSV Files")
    if folder_path:
        folder_var.set(folder_path)
        folder_entry.xview_moveto(1)  # Scroll to the end of the entry

def browse_file(index):
    file_path = filedialog.askopenfilename(
        title="Select a CSV file",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
    
    if file_path:
        entry_vars[index].set(file_path)
        file_entries[index].xview_moveto(1)  # Scroll to the end of the entry

def toggle_mode():
    global merging_mode
    merging_mode = mode_var.get()
    if merging_mode == "folder":
        folder_entry.config(state="normal")
        folder_button.config(state="normal")
        for entry, button in zip(file_entries, file_buttons):
            entry.config(state="disabled")
            button.config(state="disabled")
    else:
        folder_entry.config(state="disabled")
        folder_button.config(state="disabled")
        for entry, button in zip(file_entries, file_buttons):
            entry.config(state="normal")
            button.config(state="normal")

def browse_output_folder():
    output_path = filedialog.askdirectory(title="Select Output Folder")
    if output_path:
        output_folder_var.set(output_path)
        output_folder_entry.xview_moveto(1)

def create_file():
    filename = output_filename_var.get().strip()
    folder = output_folder_var.get().strip()
    if not filename or not folder:
        messagebox.showerror("Error", "Please specify both filename and folder.")
        return
    
    create.create_by_merging(get_csv_files(), folder, filename)
    messagebox.showinfo("Success", f"File will be saved as:\n{folder}/{filename}.xlsx")

def get_csv_files():
    if merging_mode == "folder":
        all_files = glob.glob(os.path.join(folder_var.get().strip(), "*.csv")) # returns list of all csv file path
    elif merging_mode == "individual":
        all_files = [var.get() for var in entry_vars if var.get().strip() != ""]
    else:
        return
    
    return all_files

# Main window
root = tk.Tk()
root.geometry("680x600")
root.resizable(False, False)
root.title("Multiple CSVs to Pivot Chart")

# Mode selection (folder vs individual)
mode_var = tk.StringVar(value="folder")
mode_frame = tk.Frame(root)
mode_frame.pack(pady=5, padx=3,anchor='w',fill="x")

tk.Radiobutton(mode_frame, text="Use Folder", variable=mode_var, value="folder", command=toggle_mode).pack(side="left", padx=5)
tk.Radiobutton(mode_frame, text="Use Individual Files", variable=mode_var, value="individual", command=toggle_mode).pack(side="left", padx=5)
merging_mode = ""

# Folder selector
folder_frame = tk.Frame(root)
folder_frame.pack(padx=3, pady=20, fill="x")
tk.Label(folder_frame, text="Folder:", width=6, anchor="w").pack(side="left")
folder_var = tk.StringVar()
folder_entry = tk.Entry(folder_frame, textvariable=folder_var, width=50, justify="right")
folder_entry.pack(side="left", padx=5)
folder_button = tk.Button(folder_frame, text="Browse", width=6, height=1, font=("Segoe UI", 8), command=browse_folder)
folder_button.pack(side="left", padx=1)

# Individual file selectors
entry_vars = [tk.StringVar() for _ in range(10)]
file_entries = []
file_buttons = []

for i in range(10):
    frame = tk.Frame(root)
    frame.pack(padx=3, pady=1, fill="x")
    tk.Label(frame, text=f"DUT {i+1:02d}:", width=6, anchor="w").pack(side="left")
    entry = tk.Entry(frame, textvariable=entry_vars[i], width=50, justify="right", state="disabled")
    entry.pack(side="left", padx=5)
    file_entries.append(entry)
    button = tk.Button(frame, text="Browse", width=6, height=1, font=("Segoe UI", 8),
                       state="disabled", command=lambda i=i: browse_file(i))
    button.pack(side="left", padx=1)
    file_buttons.append(button)

# Initialize state
toggle_mode()

# Output folder selector
output_folder_frame = tk.Frame(root)
output_folder_frame.pack(padx=3, pady=20, fill="x")
tk.Label(output_folder_frame, text="Output Folder:", width=12, anchor="w").pack(side="left")
output_folder_var = tk.StringVar()
output_folder_entry = tk.Entry(output_folder_frame, textvariable=output_folder_var, width=44, justify="left")
output_folder_entry.pack(side="left", padx=5)
tk.Button(output_folder_frame, text="Browse", width=6, height=1, font=("Segoe UI", 8),
          command=browse_output_folder).pack(side="left", padx=1)

# Output filename + Create button
output_filename_frame = tk.Frame(root)
output_filename_frame.pack(padx=3, pady=5, fill="x")
tk.Label(output_filename_frame, text="Output File:", width=12, anchor="w").pack(side="left")
output_filename_var = tk.StringVar()
output_filename_entry = tk.Entry(output_filename_frame, textvariable=output_filename_var, width=44, justify="left")
output_filename_entry.pack(side="left", padx=5)
create_button = tk.Button(output_filename_frame, text="Create", font=("Segoe UI", 10, "bold"),
                          bg="#4CAF50", fg="white", padx=10, pady=2, command=create_file)
create_button.pack(side="left", padx=1)

root.mainloop()
