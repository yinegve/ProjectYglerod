import os
from openpyxl import load_workbook
import pandas as pd

df = pd.read_excel("THE_THEATRE.xlsx")
yg = pd.DataFrame(df)

yg = yg.iloc[7:, 1:]
yg.columns = ['Description', 'Units', 'Quantity', 'Unit cost', 'Total cost', 'Unit price', 'Total price', 'Remark for customer']

yg.drop('Unit cost', axis=1, inplace=True)
yg.drop('Total cost', axis=1, inplace=True)
yg.drop('Unit price', axis=1, inplace=True)
yg.drop('Total price', axis=1, inplace=True)
yg.drop('Remark for customer', axis=1, inplace=True)
yg.columns = ['Description', 'Units', 'Quantity']
yg = yg[yg['Units'] != 'job']
yg = yg[yg['Units'] != 'Job']
yg = yg[yg['Units'].notna()]

yg['Quantity'] = pd.to_numeric(yg['Quantity'], errors='coerce')

for index, row in yg.iterrows():
    if "steel" in row['Description']:
        yg.at[index, 'Yglerod'] = 7
    if "plywood" in row['Description']:
        yg.at[index, 'Yglerod'] = 2
    if "Cut-Outs" in row['Description']:
        yg.at[index, 'Yglerod'] = 3
    if "cement" in row['Description']:
        yg.at[index, 'Yglerod'] = 8
    if "porcelain" in row['Description']:
        yg.at[index, 'Yglerod'] = 5
    if "aluminum profile" in row['Description']:
        yg.at[index, 'Yglerod'] = 6
    if "metal" in row['Description']:
        yg.at[index, 'Yglerod'] = 6
    if "LED Striplightsto floor" in row['Description']:
        yg.at[index, 'Yglerod'] = 3
    if "LED Spotlights to steps" in row['Description']:
        yg.at[index, 'Yglerod'] = 3
    if "Painting Works" in row['Description']:
        yg.at[index, 'Yglerod'] = 4
    if "Access Panels" in row['Description']:
        yg.at[index, 'Yglerod'] = 4
    if "Gypsum" in row['Description']:
        yg.at[index, 'Yglerod'] = 7
    if "acrylic panel" in row['Description']:
        yg.at[index, 'Yglerod'] = 8
    if "Stage Cladding" in row['Description']:
        yg.at[index, 'Yglerod'] = 8
    if "Chandelier" in row['Description']:
        yg.at[index, 'Yglerod'] = 3

for index, row in yg.iterrows():
    # Проверим, пусто ли значение в столбце "Quantity"
    if pd.isnull(row['Quantity']):
        # Если значение пустое, заменим его на 1
        yg.at[index, 'Quantity'] = 1
for index, row in yg.iterrows():
    if pd.isnull(row['Yglerod']):
        # Если значение пустое, заменим его на 1
        yg.at[index, 'Yglerod'] = 1

yg['FullYglerod'] = yg['Yglerod'] * yg['Quantity']

# Вычислим сумму значений в колонке "Sales"
total_sales = yg['FullYglerod'].sum()*0.001*24

yg.to_excel("THE_THEATRE_new.xlsx", index = False)
print('Углеродный след составляет ' + str(total_sales) + ' тонн CO2.')
if total_sales > 300:
    print('Углеродный след данной сметы строительногопроекта достаночно высок. Рекомендуем подобрать материалы с меньшим выбросом углеродного следа.')