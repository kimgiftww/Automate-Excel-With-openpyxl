import openpyxl as xl

# Construct an Excel workbook object
wb = xl.Workbook()

# insert data to the active worksheet
ws = wb.active

ws['A1'] = 'Cell A1'


# If the existing workbook is open, you must close the 
# Excel file for the wb object to save over the existing file
# otherwise, you will run into PermissionError
if not PermissionError:
    wb.save('test1.xlsx')
else:
    print('Workbook is open')