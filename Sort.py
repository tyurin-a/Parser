import os.path
import pandas as pd
import numpy as np
from numpy import nan
import openpyxl
from openpyxl import load_workbook

start_year = '1990'
end_year = '2019'
file_to_parse = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
resources = ['Coal and coal products', 'Oil products', 'Natural gas', 'Electricity',  'Heat']

file = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
xl = pd.ExcelFile(file)  # Загружаем spreadsheet (электронную таблицу) в объект pandas
# print(xl.sheet_names) # Печатаем названия листов в данном файле
dfs = {

}  # Словарь, в который выгружаем эксель-файл
dfs = xl.parse(sheet_name='Normal', skiprows=0)  # Парсим листы эксель-файла
# dfs.drop(dfs.columns[0:33], axis=1, inplace=True)
xl.close()  # Закрываем читаемый файл

if os.path.exists(file_to_parse):
    mode = "a"
    if_sheet_exists = "overlay"
else:
    mode = "w"
    if_sheet_exists = None
writer = pd.ExcelWriter(file_to_parse, engine='openpyxl', mode=mode,
                        if_sheet_exists=if_sheet_exists)  # Указываем writer библиотеки

for k in range(1, 7, 1):
    df = dfs.loc[dfs['Группа'] == k]
    if k == 1:
        i = 0
    else:
        df1 = dfs.loc[dfs['Группа'] == k - 1]
        i += (df1[df1.columns[0]].count() + 3)
    df.to_excel(writer, sheet_name='Sort', index=False, startcol=0, startrow=i)

writer.save()


