import openpyxl, os

recipesData = {}

"""for menu in os.listdir():
    if menu.startswith('меню'):
        menuWb = openpyxl.load_workbook(f'{menu}.xlsx')
        menuSheet = menuWb.active

        total_menu_weight = 0  # recipeSheet.cell(row=2, column=2).value
        total_menu_cost = 0
        total_menu_caloriesReal = 0

        for menu_recipe in range(5, menuSheet.max_row + 1):
            menu_recipe_portion_weight = menuSheet.cell(row=menu_recipe, column=2).value
            menu_recipe_price = productsData[recipeSheet.cell(row=recipe_product, column=1).value]['price']
            menu_recipe_calories100 = productsData[recipeSheet.cell(row=recipe_product, column=1).value]['calories']

            recipe_product_cost = recipe_product_price * recipe_product_weight / 1000
            recipe_product_caloriesReal = recipe_product_calories100 * recipe_product_weight / 100

            total_recipe_weight += recipeSheet.cell(row=recipe_product, column=2).value
            total_recipe_cost += recipe_product_cost
            total_recipe_caloriesReal += recipe_product_caloriesReal

            recipe_product_weight_percent = f'=(B{recipe_product}/B2)*100'
            recipe_product_cost_percent = f'=(C{recipe_product}/C2)*100'
            recipe_product_caloriesReal_percent = f'=(D{recipe_product}/D2)*100'

            recipeSheet.cell(row=recipe_product, column=5).value = recipe_product_price
            recipeSheet.cell(row=recipe_product, column=6).value = recipe_product_calories100
            recipeSheet.cell(row=recipe_product, column=3).value = recipe_product_cost
            recipeSheet.cell(row=recipe_product, column=4).value = recipe_product_caloriesReal

            recipeSheet.cell(row=2, column=2).value = total_recipe_weight
            recipeSheet.cell(row=2, column=3).value = total_recipe_cost
            recipeSheet.cell(row=2, column=4).value = total_recipe_caloriesReal

            recipeSheet.cell(row=recipe_product, column=7).value = recipe_product_weight_percent
            recipeSheet.cell(row=recipe_product, column=8).value = recipe_product_cost_percent
            recipeSheet.cell(row=recipe_product, column=9).value = recipe_product_caloriesReal_percent

        menuWb.save(menu)"""

"""menuWb = openpyxl.load_workbook('меню_1.xlsx')
menuSheet = menuWb.active
# menuProductsData = {'картопля': 2, 'капуста': 3, ...}}
menuProductsData = {}

# load one by one menu recipes sheets
for menu_recipes in range(9, menuSheet.max_row + 1):
    recipe_name = f'{menuSheet.cell(column=1, row=menu_recipes).value}.xlsx'
    recipeWb = openpyxl.load_workbook(recipe_name)
    recipeSheet = recipeWb.active

    # load and add recipes items to menuData
    for recipe_product in range(4, recipeSheet.max_row + 1):
        menuProductsData.setdefault(recipeSheet.cell(column=1, row=recipe_product).value)
    for

print(pprint.pformat(menuProductsData))"""

