import pandas as pd
import glob
import os
from pandas.io.formats import excel
from openpyxl import load_workbook
import win32com.client
import win32gui
import time
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def load_headers_from_file(folder, filename):
    merged_file = os.path.join(folder, filename + ".xlsx")
    print("merged_file:", merged_file)
    df = pd.read_excel(merged_file, sheet_name="Raw Data", header=None, nrows=5)  # just read the header row
    print("df:", df)
    print("df columns:", df.columns.tolist())

load_headers_from_file(r"C:\Users\Harvey\Desktop\Projects\csv_to_pivotChart", "m7")