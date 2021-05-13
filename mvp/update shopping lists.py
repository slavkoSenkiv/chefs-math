import openpyxl
import productsData
import menusData
import os
from openpyxl.styles import Font
bold = Font(bold=True)

shoppingListData = {}

for menu in os.listdir():
    if menu.startswith('меню'):
        menuShoppingList = openpyxl.Workbook()
        menuShoppingListSheet = menuShoppingList.active

        menuShoppingListSheet['A1'] = 'назва меню'
        menuShoppingListSheet['B1'] = 'заг кть'
        menuShoppingListSheet['C1'] = 'заг варт'
        menuShoppingListSheet['A3'] = 'продукт'
        menuShoppingListSheet['B3'] = 'к-ть'
        menuShoppingListSheet['C3'] = 'вартість'
        menuShoppingListSheet['D3'] = 'ціна'
        menuShoppingListSheet['E3'] = '% к-ті'
        menuShoppingListSheet['F3'] = '% ціни'
        menuShoppingListSheet['G3'] = '% вартості'
        menuShoppingListSheet['A2'] = menu[:-5]

        menuShoppingListSheet['A1'].font = bold
        menuShoppingListSheet['B1'].font = bold
        menuShoppingListSheet['C1'].font = bold
        menuShoppingListSheet['A3'].font = bold
        menuShoppingListSheet['B3'].font = bold
        menuShoppingListSheet['C3'].font = bold
        menuShoppingListSheet['D3'].font = bold
        menuShoppingListSheet['E3'].font = bold
        menuShoppingListSheet['F3'].font = bold
        menuShoppingListSheet['G3'].font = bold

        menuWb = openpyxl.load_workbook(menu)
        menuWbSheet = menuWb.active

        for menuRecipes in range(5, menuWbSheet.max_row + 1):
            if menuWbSheet.cell(row=menuRecipes, column=1).value in menusData.menusData[menu]:
                print(menuWbSheet.cell(row=menuRecipes, column=1).value)








        menuShoppingList.save('тест шопінг  ' + menu)

