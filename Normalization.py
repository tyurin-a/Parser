# Импорт библиотек
import os.path
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

start_year = '1990'
end_year = '2019'
file_to_parse = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
resources = ['Coal and coal products', 'Oil products', 'Natural gas', 'Electricity',  'Heat']

file = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
xl = pd.ExcelFile(file)  # Загружаем spreadsheet (электронную таблицу) в объект pandas
# print(xl.sheet_names) # Печатаем названия листов в данном файле
df = []
df = xl.parse(sheet_name='Total', skiprows=0)  # Парсим листы эксель-файла
xl.close()  # Закрываем читаемый файл

if os.path.exists(file_to_parse):
    mode = "a"
    if_sheet_exists = "overlay"
else:
    mode = "w"
    if_sheet_exists = None
writer = pd.ExcelWriter(file_to_parse, engine='openpyxl', mode=mode,
                        if_sheet_exists=if_sheet_exists)  # Указываем writer библиотеки

df.reset_index(drop=True)  # Cбрасываем индексирование строк
df['Max'] = df.max(axis=1, numeric_only=True)  # Находим максимальное значение в каждой строке среди численных данных
df1 = df['Max']
country = df['COUNTRY']
df2 = df.loc[:, start_year:end_year]
df1.to_excel(writer, sheet_name='Total', index=False, startcol=31)  # Записываем столбец с максимальным значением в файл.
country.to_excel(writer, sheet_name='Total', index=False, startcol=33) # Записываем столбец стран для нормированной таблицы.

# Нормируем таблицу на максимальное значение
for i in range(0, 160):
    val = df2.iloc[i]  # Выбираем строку значений из df2
    max_val = df1.iloc[i] # Выбираем строку значений из df1 (максимальное значение)
    df2.iloc[i] = val / max_val
    df2.to_excel(writer, sheet_name='Total', index=False, startcol=34)

writer.save()  # Сохраняем результат