import pandas as pd
import glob
import os
from pandas.io.formats import excel
from openpyxl import load_workbook

excel.ExcelFormatter.header_style = None # remove the formatting in header

use_case = input("Enter the user case: 'create' or 'replace' ").lower()
if use_case == "create":
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

    ### CREATE
    # Read and merge CSV files
    df_list = [pd.read_csv(f, parse_dates=False) for f in all_files]
    merged_df = pd.concat(df_list, ignore_index=True)
    # Write to one Excel file
    merged_df.to_excel(output_file, index=False, sheet_name="Raw Data")
    print(f"Merged {len(all_files)} files into {output_file}")  

elif use_case== "replace":
    ### REPLACE
    merged_path = input("Enter the merged path (should already exist): ")
    unit_to_update = input("Enter the unit name to update: ")
    new_csv_path = input("Enter the csv containing the new data: ")

    # --- Load Data ---
    # Load the new CSV containing updated data for Unit_2
    new_data = pd.read_csv(new_csv_path)

    # Load the existing Excel Raw Data sheet
    existing_df = pd.read_excel(merged_path, sheet_name="Raw Data")

    # --- Replace the rows for the specific unit ---
    # Drop old rows for this unit
    filtered_df = existing_df[existing_df["Unit name"] != unit_to_update]

    # Append the new data for this unit
    updated_df = pd.concat([filtered_df, new_data], ignore_index=True)

    # --- Write back to Excel ---
    with pd.ExcelWriter(
        merged_path,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        updated_df.to_excel(writer, index=False, sheet_name="Raw Data")

    print(f"✅ Successfully replaced data for {unit_to_update} in Raw Data sheet.")
    
else:
    print("Use case selected is not supported. Exiting...")
    quit()