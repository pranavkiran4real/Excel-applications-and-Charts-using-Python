
import openpyxl

import random

# import AreaChart3D class from openpyxl.chart sub_module
from openpyxl.chart import AreaChart3D,Reference

wb = openpyxl.Workbook()
sheet = wb.active


"""
# write 0 to 9 in 1st column of the active sheet
for i in range(9):
    sheet.append([random.randint(1,100)])
"""

sheet.append([587])
sheet.append([62])
sheet.append([234])
sheet.append([4523])
sheet.append([56])
sheet.append([3253])
sheet.append([145])
sheet.append([2233])
sheet.append([20])


values = Reference(sheet, min_col = 1, min_row = 1,
                         max_col = 1, max_row = 8)

# Create object of AreaChart3D class
chart = AreaChart3D()

chart.add_data(values)

# set the title of the chart
chart.title = " AREA-CHART3D "

# set the title of the x-axis
chart.x_axis.title = " X-AXIS "

# set the title of the y-axis
chart.y_axis.title = " Y-AXIS "

# add chart to the sheet
# the top-left corner of a chart
# is anchored to cell E2 .
sheet.add_chart(chart, "E2")

# save the file
wb.save("AreaChart3D.xlsx")
