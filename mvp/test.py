import openpyxl

wb = openpyxl.load_workbook('text.xlsx')
sheet = wb.active

dic = {'inside': {'x': 1.5, 'y': 2.0}}

a = 1
for letters in dic['inside']:

    sheet[f'A{a}'] = letters
    sheet[f'B{a}'] = dic['inside'][letters]
    a += 1

wb.save('text.xlsx')
