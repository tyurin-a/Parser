# Импорт библиотек
import os.path
import pandas as pd
import openpyxl
from openpyxl import load_workbook

start_year = '1990'
end_year = '2019'
file_to_parse = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'

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
    table1 = df.loc[:, 'COUNTRY':'Useful consumption']  # Создаем датафреймы из двух таблиц на листе
    table2 = df.loc[:, 'Country Name':'Value added']
    table1 = pd.merge(table1, table2, left_on=['COUNTRY'], right_on=['Country Name'], how='left')  # Аналог ВПР
    table1.drop(['Country Name'], axis='columns', inplace=True)  # Удаляем лишний столбец
    # print(df1.name)  # Печатаем названия первичных ключей (названия столбцов) в данном массиве (не в датафрейме)
    # print(df.keys())  # Печатаем названия первичных ключей (названия столбцов) в данном датафрейме
    table1.to_excel(writer, sheet_name=str(k), index=False, startcol=0)  # Записываем датафрейм в файл.
# index=False отключает запись индексов, startcol=1 начианет запись с 1 стобца (нумерация с нуля).
writer.save()  # Сохраняем результат