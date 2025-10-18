import pandas as pd
import glob
import os

merging_mode_input = input("Enter merging mode 'group' or 'individual': ").lower()
merging_mode = True # default, group
if merging_mode_input == 'group':
    merging_mode = True
elif merging_mode_input == 'individual':
    merging_mode = False
else:
    print("Mode selected is not supported. Exiting...")
    quit()

folder = ""
all_files = list()

if merging_mode:
    folder = input("Enter the folder containing all csv files: ")
    all_files = glob.glob(os.path.join(folder, "*.csv")) # returns list of all csv file path
else:
    for i in range(3):
        csv_path = input("Enter the filename of csv" + str(i+1) + ": ")
        all_files.append(csv_path)

output_folder = input("Enter the folder to save output: ")
output_filename = input("Enter output filename: ")
output_file = os.path.join(output_folder, output_filename + ".xlsx") # Folder containing all CSV files

merged_df = pd.DataFrame()

# Read and merge CSV files
df_list = [pd.read_csv(f, parse_dates=False) for f in all_files]
merged_df = pd.concat(df_list, ignore_index=True)

# Write to one Excel file
merged_df.to_excel(output_file, index=False, sheet_name="Raw Data")

print(f"Merged {len(all_files)} files into {output_file}")  
