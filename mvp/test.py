import openpyxl

wb = openpyxl.load_workbook('меню 3.xlsx')
for sheets in wb.sheetnames:
    print(sheets)