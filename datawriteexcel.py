import xlwt
from xlwt import Workbook

wb=Workbook()

sheet1=wb.add_sheet("Sheet2")

count=int(input("Enter Count Records : "))

j=0

for i in range(count):
    out=input("Enter Content : ")
    sheet1.write(i,j,out)

wb.save('xlwt example.xls')