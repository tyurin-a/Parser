import os.path
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

start_year = '1990'
end_year = '2019'
file_to_parse = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
resources = ['Coal and coal products', 'Oil products', 'Natural gas', 'Electricity',  'Heat']

# Задаем стартовые параметры для парсинга и фильтрации (начальный и конечный года, энергоресурсы, отрасль и путь к файлу)
file = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
xl = pd.ExcelFile(file)  # Загружаем spreadsheet (электронную таблицу) в объект pandas
# print(xl.sheet_names) # Печатаем названия листов в данном файле
dfs = {

}  # Словарь, в который выгружаем эксель-файл
for k in xl.sheet_names:
    dfs[k] = xl.parse(sheet_name=str(k), skiprows=0)  # Парсим листы эксель-файла
xl.close()  # Закрываем читаемый файл

if os.path.exists(file_to_parse):
    mode = "a"
    if_sheet_exists = "overlay"
else:
    mode = "w"
    if_sheet_exists = None
writer = pd.ExcelWriter(file_to_parse, engine='openpyxl', mode=mode,
                        if_sheet_exists=if_sheet_exists)  # Указываем writer библиотеки

dfc = dfs[start_year]
country = dfc['COUNTRY']
country.to_excel(writer, sheet_name='Total', index=False, startcol=0)
i = 1
for k in dfs:
    df = dfs[k]  # Получаем лист из словаря dfs
    df1 = df['Specific energy consumption']
    df1.name = k
    # print(df1.name)  # Печатаем названия первичных ключей (названия столбцов) в данном массиве (не в датафрейме)
    # print(df.keys())  # Печатаем названия первичных ключей (названия столбцов) в данном датафрейме
    df1.to_excel(writer, sheet_name='Total', index=False, startcol=i)  # Записываем датафрейм в файл.
# index=False отключает запись индексов, startcol=1 начианет запись с 1 стобца (нумерация с нуля).
    i += 1
writer.save()  # Сохраняем результат