import openpyxl, os
from pathlib import Path
import pprint

# calculate each recipe item cost based on product sheet
productWb = openpyxl.load_workbook('список продуктів.xlsx')
productSheet = productWb.active

# update product properties data
# productsData = {'product_name': {'price': product_price, 'calories': product_calories}}
productsData = {}
for products in range(2, productSheet.max_row + 1):
    product_name     = productSheet['A' + str(products)].value
    product_price    = productSheet['B' + str(products)].value
    product_calories = productSheet['C' + str(products)].value

    # make sure the key for this product exists
    productsData.setdefault(product_name, {'price': 0, 'calories': 0})

    productsData[product_name]['price'] = product_price
    productsData[product_name]['calories'] = product_calories

for recipe in os.listdir(Path.cwd()):
    if recipe.startswith('рецепт'):

        recipeWb = openpyxl.load_workbook(recipe)
        recipeSheet = recipeWb.active

        total_recipe_weight = 0  # recipeSheet.cell(row=2, column=2).value
        total_recipe_cost = 0
        total_recipe_caloriesReal = 0

        for recipe_product in range(4, recipeSheet.max_row + 1):
            recipe_product_weight = recipeSheet.cell(row=recipe_product, column=2).value
            recipe_product_price = productsData[recipeSheet.cell(row=recipe_product, column=1).value]['price']
            recipe_product_calories100 = productsData[recipeSheet.cell(row=recipe_product, column=1).value]['calories']

            recipe_product_cost  = recipe_product_price * recipe_product_weight / 1000
            recipe_product_caloriesReal  = recipe_product_calories100 * recipe_product_weight / 100

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

        recipeWb.save(recipe)

productsDataDoc = open('productsData.py', 'w', encoding='utf-8')
productsDataDoc.write('productsData = ' + pprint.pformat(productsData))
productsDataDoc.close()