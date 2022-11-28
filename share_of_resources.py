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
file_to_parse = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
resources = ['Coal and coal products', 'Oil products', 'Natural gas', 'Electricity',  'Heat']

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

for k in dfs:
    df = dfs[k]  # Получаем лист из словаря dfs
    df['Coal and coal products (%)'] = df['Coal and coal products'] * 0.35 / df['Useful consumption']
    df['Oil products (%)'] = df['Oil products'] * 0.35 / df['Useful consumption']
    df['Natural gas (%)'] = df['Natural gas'] * 0.35 / df['Useful consumption']
    df['Electricity (%)'] = df['Electricity'] * 0.9 / df['Useful consumption']
    df['Heat (%)'] = df['Heat'] * 0.9 / df['Useful consumption']
    # print(df1.name)  # Печатаем названия первичных ключей (названия столбцов) в данном массиве (не в датафрейме)
    # print(df.keys())  # Печатаем названия первичных ключей (названия столбцов) в данном датафрейме
    df1 = df[['Coal and coal products (%)', 'Oil products (%)', 'Natural gas (%)', 'Electricity (%)', 'Heat (%)']]
    df1.to_excel(writer, sheet_name=str(k), index=False, startcol=13)  # Записываем датафрейм в файл.
    if k == '2019':
        break
# index=False отключает запись индексов, startcol=1 начианет запись с 1 стобца (нумерация с нуля).
writer.save()  # Сохраняем результат