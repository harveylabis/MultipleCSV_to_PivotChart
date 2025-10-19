import pandas as pd
import glob
import os
from pandas.io.formats import excel
from openpyxl import load_workbook

excel.ExcelFormatter.header_style = None # remove the formatting in header

def create_by_merging(all_files, output_folder, output_filename):    
    output_file = os.path.join(output_folder, output_filename + ".xlsx") # Folder containing all CSV files
    merged_df = pd.DataFrame()

    # Read and merge CSV files
    df_list = [pd.read_csv(f, parse_dates=False) for f in all_files]
    merged_df = pd.concat(df_list, ignore_index=True)
    # Write to one Excel file
    merged_df.to_excel(output_file, index=False, sheet_name="Raw Data")
    
    print(f"Merged {len(all_files)} files into {output_file}")  