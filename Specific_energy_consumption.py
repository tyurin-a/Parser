# Импорт библиотек
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook

# Задаем стартовые параметры для парсинга и фильтрации (начальный и конечный года, энергоресурсы, отрасль и путь к файлу)
start_year = '1990'
end_year = '2019'
file = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
file_to_parse = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
xl = pd.ExcelFile(file)  # Загружаем spreadsheet (электронную таблицу) в объект pandas
# print(xl.sheet_names) # Печатаем названия листов в данном файле
dfs = {

}  # Словарь, в который выгружаем эксель-файл
for k in xl.sheet_names:
    dfs[k] = xl.parse(sheet_name=str(k), skiprows=0)  # Парсим листы эксель-файла
xl.close()  # Закрываем читаемый файл
writer = pd.ExcelWriter(file_to_parse, engine='openpyxl', mode="a",
                        if_sheet_exists="overlay")  # Указываем writer библиотеки
for k in dfs:
    df = dfs[k]  # Получаем лист из словаря dfs
    table1 = df.loc[:, 'COUNTRY':'Value added']  # Создаем датафрейм
    table1['Specific energy consumption'] = (table1['Useful consumption']/table1['Value added'])
    # print(df1.name)  # Печатаем названия первичных ключей (названия столбцов) в данном массиве (не в датафрейме)
    # print(df.keys())  # Печатаем названия первичных ключей (названия столбцов) в данном датафрейме
    table1.to_excel(writer, sheet_name=str(k), index=False, startcol=0)  # Записываем датафрейм в файл.
# index=False отключает запись индексов, startcol=1 начианет запись с 1 стобца (нумерация с нуля).
writer.save()  # Сохраняем результат

