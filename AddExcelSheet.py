#Adding style sheet in excel
# importing xlwt module

#pip install xlwt

import xlwt

workbook = xlwt.Workbook()

sheet = workbook.add_sheet("Sheet Name")

# Specifying style
style = xlwt.easyxf('font: bold 1')

# Specifying column
sheet.write(0, 0, 'File', style)
workbook.save("example.xls")
