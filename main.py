import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import create
import glob
import os
import pandas as pd

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

def load_headers_from_file(merged_file):
    df = pd.read_excel(merged_file, nrows=1)  # just read the header row
    header_listbox.delete(0, tk.END)
    for col in df.columns:
        header_listbox.insert(tk.END, col)

# Main window
root = tk.Tk()
root.geometry("1920x1080")
# root.resizable(False, False)
root.title("Multiple CSVs to Pivot Chart")

left_frame = tk.Frame(root, width=700, height=650, bg='lightblue')
left_frame.pack_propagate(False) # False - use the defined size
left_frame.pack(side="left", padx=10, pady=10)

# Frame for mode
mode_frame = ttk.Frame(left_frame, width=700, height=50, borderwidth=10, relief=tk.GROOVE)
mode_frame.pack_propagate(False) # False - use the defined size
mode_frame.pack(side="top", padx=10, pady=10) 

# Mode selection (folder vs individual)
mode_var = tk.StringVar(value="folder")
ttk.Label(mode_frame, text="Merging Mode:").pack(side="left", padx=5)
ttk.Radiobutton(mode_frame, text="Use Folder", variable=mode_var, value="folder", command=toggle_mode).pack(side="left", padx=10)
ttk.Radiobutton(mode_frame, text="Use Individual Files", variable=mode_var, value="individual", command=toggle_mode).pack(side="left", padx=5)

# Frame for folder
folder_frame = ttk.Frame(left_frame, width=700, height=50, borderwidth=10, relief=tk.GROOVE)
folder_frame.pack_propagate(False) # False - use the defined size
folder_frame.pack(side="top", padx=10, pady=10) 

# Folder selector
ttk.Label(folder_frame, text="Folder:", width=6).pack(side="left")
folder_var = tk.StringVar()
folder_entry = ttk.Entry(folder_frame, textvariable=folder_var, width=50, justify="right")
folder_entry.pack(side="left", padx=5)
folder_button = ttk.Button(folder_frame, text="Browse", width=6, command=browse_folder)
folder_button.pack(side="left", padx=1)

# Frame for individual csv files inputs
individual_csv_frame = ttk.Frame(left_frame, width=700, height=350, borderwidth=10, relief=tk.GROOVE)
individual_csv_frame.pack_propagate(False) # False - use the defined size
individual_csv_frame.pack(side="top", padx=10, pady=10) 

# Individual file selectors
entry_vars = [tk.StringVar() for _ in range(10)]
file_entries = []
file_buttons = []

for i in range(10):
    frame = ttk.Frame(individual_csv_frame)
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

# Frame for output folder
output_folder_frame = ttk.Frame(left_frame, width=700, height=50, borderwidth=10, relief=tk.GROOVE)
output_folder_frame.pack_propagate(False) # False - use the defined size
output_folder_frame.pack(side="top", padx=10, pady=10) 

# Output folder
ttk.Label(output_folder_frame, text="Output Folder:", width=12, anchor="w").pack(side="left")
output_folder_var = tk.StringVar()
output_folder_entry = ttk.Entry(output_folder_frame, textvariable=output_folder_var, width=44, justify="left")
output_folder_entry.pack(side="left", padx=5)
tk.Button(output_folder_frame, text="Browse", width=6, height=1, font=("Segoe UI", 8),
          command=browse_output_folder).pack(side="left", padx=1)

# Frame for output folder
output_filename_frame = ttk.Frame(left_frame, width=700, height=50, borderwidth=10, relief=tk.GROOVE)
output_filename_frame.pack_propagate(False) # False - use the defined size
output_filename_frame.pack(side="top", padx=10, pady=10) 

# Output filename + Create button
ttk.Label(output_filename_frame, text="Output name:", width=12, anchor="w").pack(side="left")
output_filename_var = tk.StringVar()
output_filename_entry = ttk.Entry(output_filename_frame, textvariable=output_filename_var, width=44, justify="left")
output_filename_entry.pack(side="left", padx=5)
create_button = tk.Button(output_filename_frame, text="Create", font=("Segoe UI", 10, "bold"),
                          bg="#4CAF50", fg="white", padx=10, pady=2, command=create_file)
create_button.pack(side="left", padx=1)

# CENTER FRAME 
center_frame = tk.Frame(root, width=700, height=650, borderwidth=2, relief=tk.GROOVE)
center_frame.pack_propagate(False) # False - use the defined size
center_frame.pack(side="left", padx=25, pady=10)

# Label
tk.Label(center_frame, text="Headers: ").pack(anchor="w", padx=5)

# Scrollable Listbox for headers
header_frame = ttk.Frame(center_frame)
header_frame.pack(fill="both", expand=True, padx=5)
header_listbox = tk.Listbox(header_frame, selectmode="extended", width=25, height=10)
header_listbox.pack(side="left", fill="both", expand=True, padx=5)
header_scrollbar = tk.Scrollbar(header_frame, orient="vertical", command=header_listbox.yview)
header_scrollbar.pack(side="right", fill="y")
header_listbox.config(yscrollcommand=header_scrollbar.set)

# ---- Axis, Legend, Values ----
# Frame for axis
axis_frame = ttk.Frame(center_frame, width=700, height=50, borderwidth=10, relief=tk.GROOVE)
axis_frame.pack_propagate(False) # False - use the defined size
axis_frame.pack(side="top", padx=5, pady=5) 
ttk.Label(axis_frame, text="Axis:").pack(side='left', padx=5)
axis_var = tk.StringVar()
ttk.Entry(axis_frame, textvariable=axis_var, width=25).pack(side='left', padx=30)

# Frame for legend
legend_frame = ttk.Frame(center_frame, width=700, height=50, borderwidth=10, relief=tk.GROOVE)
legend_frame.pack_propagate(False) # False - use the defined size
legend_frame.pack(side="top", padx=5, pady=5) 
ttk.Label(legend_frame, text="Legend:").pack(side='left', padx=5)
axis_var = tk.StringVar()
ttk.Entry(legend_frame, textvariable=axis_var, width=25).pack(side='left', padx=5)

# Frame for values
value_frame = ttk.Frame(center_frame, width=700, height=50, borderwidth=10, relief=tk.GROOVE)
value_frame.pack_propagate(False) # False - use the defined size
value_frame.pack(side="top", padx=5, pady=5) 
ttk.Label(value_frame, text="Value:").pack(side='left', padx=5)
axis_var = tk.StringVar()
ttk.Entry(value_frame, textvariable=axis_var, width=25).pack(side='left', padx=18)

# ---- Create Pivot Chart Button ----
# Frame for output folder
generate_pc_frame = ttk.Frame(center_frame, width=700, height=50, borderwidth=10, relief=tk.GROOVE)
generate_pc_frame.pack_propagate(False) # False - use the defined size
generate_pc_frame.pack(side="top", padx=5, pady=5) 

# Output pivotchart name + Create button
ttk.Label(generate_pc_frame, text="PivotChart name:", width=15, anchor="w").pack(side="left")
pivotChart_var = tk.StringVar()
pivotChart_entry = ttk.Entry(generate_pc_frame, textvariable=output_filename_var, width=25, justify="left")
pivotChart_entry.pack(side="left", padx=5)
tk.Button(generate_pc_frame, text="Generate PivotChart", width=25, height=2, font=("Segoe UI", 9, "bold"), bg="#4CAF50",fg="white").pack(side='left')

load_headers_from_file(r'C:\Users\Harvey\Desktop\Projects\csv_to_pivotChart\Merged_Data.xlsx')

root.mainloop()
