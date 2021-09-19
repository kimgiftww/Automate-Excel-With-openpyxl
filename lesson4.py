import datetime as dt
import openpyxl as xl

"""
Working with an existing Excel workbook
"""
wb = xl.load_workbook('data1.xlsm')
print(wb.sheetnames)

wsData1 = wb['Data1']
for row in wsData1:
    for cell in row:
        print(cell.value)

# save existing workbook as a new file
wb.save('data1_modified.xlsx')