import openpyxl
from openpyxl import Workbook

wb=Workbook()

#sheet1=wb.add_sheet("Sheet2")
sheet1=wb.active

#sheet1.write(1,1,"First Data")
count=int(input("Enter Count Records : "))

j=1

for i in range(count):
    if(i!=0):
        out=input("Enter Content : ")
        sheet1.cell(i,j,out)


wb.save('xlwt example.xlsx')