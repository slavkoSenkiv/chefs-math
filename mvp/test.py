import openpyxl

testWb = openpyxl.load_workbook('test.xlsx')
testWbSheet = testWb.active

for rows in range(4, testWbSheet.max_row + 1):
    testWbSheet.cell(row=rows, column=2).value = 'a'

for rows in range(4, testWbSheet.max_row):
    testWbSheet.cell(row=rows, column=3).value = 'b'

testWb.save('test.xlsx')
