import win32com.client
import win32gui
import time


# Paths to your files
updater_path = r"C:\Users\Harvey\Desktop\Projects\csv_to_pivotChart\Updater.xlsm"
merged_path = r"C:\Users\Harvey\Desktop\Projects\csv_to_pivotChart\Merged_Data.xlsx"

# Start Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True  # Set to True to automatically open the merged file after macro runs

# Open the Updater file
updater_wb = excel.Workbooks.Open(updater_path)

# Run your macro
# excel.Application.Run("Module2.CreatePivotChartInMergedData")
pc_axis = "DateTime"
pc_legend = "Unit name"
pc_values = "Voltage"
pivotName = pc_axis + " over " + pc_values

excel.Application.Run("Module3.CreatePivotChartInMergedDataVariable", merged_path, pivotName, pc_axis, pc_legend, pc_values)

# Save & close Updater (optional)
updater_wb.Save()
updater_wb.Close()

# Open the merged data file
merged_wb = excel.Workbooks.Open(merged_path)
print(f"Opened merged file: {merged_wb.Name}")

# (Optional) — you can activate a specific sheet like PivotChart:
merged_wb.Sheets(pivotName).Activate()

# Give Excel some time to fully open
time.sleep(2)

# Bring Excel to the front
win32gui.SetForegroundWindow(win32gui.FindWindow(None, excel.Caption))

# Excel will stay open for viewing.
# If you want to close later in the script, do:
# merged_wb.Close(SaveChanges=False)
# excel.Quit()


