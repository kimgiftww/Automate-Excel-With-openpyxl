# Getting-Started-With-Openpyxl

openpyxl is a Python library to read/write Excel xlsx/xlsm/xltx/xltm files. It was born from lack of existing library to read/write natively from Python the Office Open XML format. openpyxl is probably the most widely use Python library to manage Excel spreadsheet today, even pandas uses openpyxl as the default engine to work with Excel application.

**Pros**
- [x] Probably the most popular Excel package for read/write data. pandas uses openpyxl as the default Excel engine to read/write data.
- [x] Accepts DataTime object. 

**Cons**
- [x] Cannot interact with Excel workbook while file is open (this is possible with pywin32).
- [x] Cannot read a formula cell.
- [x] Cannot access Excel object models like pywin32. 
