import win32com.client

# Path to your xlsm file
file_path = r"C:\Users\Harvey\Desktop\Projects\csv_to_pivotChart\Updater.xlsm"

# Start Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # or True if you want to see Excel open


# Open the workbook
wb = excel.Workbooks.Open(file_path)

# Run the macro (example: Macro1 in Module1)
excel.Application.Run("Module1.AutoUpdateMergedData")

# Save & close
wb.Save()
wb.Close()

# Quit Excel
#excel.Quit()
