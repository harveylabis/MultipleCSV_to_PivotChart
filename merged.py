import pandas as pd
import glob
import os

# Folder containing all CSV files
folder = r"C:\Users\Harvey\Desktop\Projects\csv_to_pivotChart"
output_file = os.path.join(folder, "Merged_Data.xlsx")

# Find all CSV files in the folder
all_files = glob.glob(os.path.join(folder, "*.csv")) # returns list of all csv file path

# Read and merge CSV files
df_list = [pd.read_csv(f) for f in all_files]
merged_df = pd.concat(df_list, ignore_index=True)

# Write to one Excel file
merged_df.to_excel(output_file, index=False, sheet_name="Raw Data")

print(f"Merged {len(all_files)} files into {output_file}")