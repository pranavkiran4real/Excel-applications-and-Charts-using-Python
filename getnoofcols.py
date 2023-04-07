# importing openpyxl module
import openpyxl

# Give the location of the file
path = "D:\\demo.xlsx"

# workbook object is created
wb_obj = openpyxl.load_workbook(path)

sheet_obj = wb_obj.active

# ptint total number of column
print(sheet_obj.max_column)
