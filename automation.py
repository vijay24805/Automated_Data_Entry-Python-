import xlrd
import re
file_location = "D:/Entry Automation/Automated_Data_Entry-Python-/2014_ERsBBL.xlsx"
workbook = xlrd.open_workbook(file_location)

sheet = workbook.sheet_by_index(0)
sheet.cell_value(0,0)
print sheet.nrows
sheet.ncols

for row in range(10588,10940):
    sheet.cell_value(row,2)


data =[[sheet.cell_value(r,c) for c in range(0,4)] for r in range(10587,10940)]
print len(data)

m = re.split('\s+',data[300][2])
print m


k=0
for i in range(0,len(data)):
    for j in range(2,4):
        data[i][j]
    


