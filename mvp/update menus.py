import openpyxl, os, pprint
import productsData

"""
menusData = {'menu': {'menuRecipes': {'menuRecipe1': {'totalRecipeData': {'totalRecipeWeight': 0, 'totalRecipePrice': 0, 'totalRecipeCalories': 0},
                                                      'recipeProducts':  {'menuRecipeProduct1Weight': 0, 'menuRecipeProduct2Weight': 0}}}
"""

menusData = {}

for menu in os.listdir():
    if menu.startswith('меню'):
        menusData.setdefault(menu, {'menuRecipes': {'totalRecipeData': {'totalRecipeWeight': 0, 'totalRecipePrice': 0, 'totalRecipeCalories': 0},
                                                      'recipeProducts':  {}}})
        menuWb = openpyxl.load_workbook(menu)
        menuWbSheet = menuWb.active

        menuWeightPerPerson   = 0

        for menu_recipes in range(5, menuWbSheet.max_row + 1):

            menuWeightPerPerson += menuWbSheet.cell(row=menu_recipes, column=2).value

            menuRecipeName = menuWbSheet.cell(row=menu_recipes, column=1).value
            menuRecipeWb = openpyxl.load_workbook(menuRecipeName + '.xlsx')
            menuRecipeWbSheet = menuRecipeWb.active

            menusData[menu]['menuRecipes']['totalRecipeData']['totalRecipeWeight'] = menuRecipeWbSheet.cell(row=2, column=2).value
            menusData[menu]['menuRecipes']['totalRecipeData']['totalRecipePrice'] = menuRecipeWbSheet.cell(row=2, column=3).value
            menusData[menu]['menuRecipes']['totalRecipeData']['totalRecipeCalories'] = menuRecipeWbSheet.cell(row=2, column=4).value

            for recipeProducts in range(4, menuRecipeWbSheet.max_row + 1):
                menusData[menu]['menuRecipes']['recipeProducts'][menuRecipeWbSheet.cell(row=recipeProducts, column=1).value] = menuRecipeWbSheet.cell(row=recipeProducts, column=2).value

        menuWbSheet.cell(row=3, column=2).value = menuWeightPerPerson

        print(menuWeightPerPerson)

print(pprint.pformat(menusData))


