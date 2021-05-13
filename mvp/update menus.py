import openpyxl, os, pprint

"""
menusData = {'menu1': {'menuRecipe1': {'product1': 0, 'product2': 0}, 
                       'menuRecipe2': {'product2': 0, 'product3': 0}}                                                     
"""

menusData = {}

for menu in os.listdir():
    if menu.startswith('меню'):
        menusData.setdefault(menu, {})
        menuWb = openpyxl.load_workbook(menu)
        menuWbSheet = menuWb.active

        menuWeightPerPerson = 0
        menuCostPerPerson = 0
        menuCaloriesPerPerson = 0

        for menu_recipes in range(5, menuWbSheet.max_row + 1):
            # increment in order to get menuWeightPerPerson
            menuRecipeName = menuWbSheet.cell(row=menu_recipes, column=1).value
            menusData[menu].setdefault(menuRecipeName, {})
            menuRecipeWb = openpyxl.load_workbook(menuRecipeName + '.xlsx')
            menuRecipeWbSheet = menuRecipeWb.active

            # recipe total weight
            recipeTotalWeight = menuRecipeWbSheet.cell(row=2, column=2).value
            # recipe price per 1kg
            menuWbSheet.cell(row=menu_recipes, column=5).value = 1000 * menuRecipeWbSheet.cell(row=2, column=3).value / recipeTotalWeight
            # recipe calories per 100
            menuWbSheet.cell(row=menu_recipes, column=6).value = 100 * menuRecipeWbSheet.cell(row=2, column=4).value / recipeTotalWeight
            # menu recipe cost per portion
            menuWbSheet.cell(row=menu_recipes, column=3).value = menuWbSheet.cell(row=menu_recipes, column=2).value * menuWbSheet.cell(row=menu_recipes, column=5).value / 1000
            # menu recipe calories per portion
            menuWbSheet.cell(row=menu_recipes, column=4).value = menuWbSheet.cell(row=menu_recipes, column=2).value * menuWbSheet.cell(row=menu_recipes, column=6).value / 100

            menuWeightPerPerson += menuWbSheet.cell(row=menu_recipes, column=2).value
            menuCostPerPerson += menuWbSheet.cell(row=menu_recipes, column=3).value
            menuCaloriesPerPerson += menuWbSheet.cell(row=menu_recipes, column=4).value

            for recipeProducts in range(4, menuRecipeWbSheet.max_row + 1):
                menusData[menu][menuRecipeName].setdefault(menuRecipeWbSheet.cell(row=recipeProducts, column=1).value, menuRecipeWbSheet.cell(row=recipeProducts, column=2).value)

        menuWbSheet.cell(row=3, column=2).value = menuWeightPerPerson
        menuWbSheet.cell(row=3, column=3).value = menuCostPerPerson
        menuWbSheet.cell(row=3, column=4).value = menuCaloriesPerPerson

        menuWbSheet.cell(row=2, column=2).value = menuWeightPerPerson * menuWbSheet.cell(row=1, column=2).value
        menuWbSheet.cell(row=2, column=3).value = menuCostPerPerson * menuWbSheet.cell(row=1, column=2).value
        menuWbSheet.cell(row=2, column=4).value = menuCaloriesPerPerson * menuWbSheet.cell(row=1, column=2).value

        menuWb.save(menu)

menusDataDoc = open('menusData.py', 'w', encoding='utf-8')
menusDataDoc.write('menusData = ' + pprint.pformat(menusData))
menusDataDoc.close()

print(pprint.pformat(menusData))


