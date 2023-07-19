# Импорт библиотек
import os.path
import pandas as pd
from tqdm import trange
import numpy as np
from numpy import nan
import openpyxl
from openpyxl import load_workbook
from openpyxl.chart import Reference, BarChart, LineChart, Series
from openpyxl.chart.label import DataLabelList

start_year = '1990'
end_year = '2019'
file_to_parse = r'C:\Users\Артем\Desktop\Residential consumption.xlsx'
resources = ['Coal and coal products', 'Oil products', 'Natural gas', 'Biofuels and waste', 'Electricity',  'Heat']

file_group = r'C:\Users\Артем\Desktop\Group_keys_residential.xlsx'
xl = pd.ExcelFile(file_group)  # Загружаем spreadsheet (электронную таблицу) в объект pandas
# print(xl.sheet_names) # Печатаем названия листов в данном файле
df_group = {

}

df_group = xl.parse(skiprows=0)  # Парсим лист эксель-файла
xl.close()  # Закрываем читаемый файл

table_group = df_group.loc[:, 'Country':'Group']

file = r'C:\Users\Артем\Desktop\Residential consumption.xlsx'
xl = pd.ExcelFile(file)  # Загружаем spreadsheet (электронную таблицу) в объект pandas
# print(xl.sheet_names) # Печатаем названия листов в данном файле
dfs = {

}

dfs = xl.parse(sheet_name='Normal', skiprows=0)  # Парсим лист эксель-файла

xl.close()  # Закрываем читаемый файл

if os.path.exists(file_to_parse):
    mode = "a"
    if_sheet_exists = "overlay"
else:
    mode = "w"
    if_sheet_exists = None
writer = pd.ExcelWriter(file_to_parse, engine='openpyxl', mode=mode,
                        if_sheet_exists=if_sheet_exists)  # Указываем writer библиотеки

df = dfs  # Получаем лист из словаря dfs
table = df.loc[:, 'COUNTRY':'2019']
table = pd.merge(table, table_group, left_on=['COUNTRY'], right_on=['Country'],
                 how='left')
table.drop(['Country'], axis='columns', inplace=True)
# print(df1.name)  # Печатаем названия первичных ключей (названия столбцов) в данном массиве (не в датафрейме)
# print(df.keys())  # Печатаем названия первичных ключей (названия столбцов) в данном датафрейме
table.to_excel(writer, sheet_name='Normal', index=False, startcol=0)
# index=False отключает запись индексов, startcol=1 начианет запись с 1 стобца (нумерация с нуля).
writer.save()  # Сохраняем результат