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

industry_dict = {"China": "People's Republic of China", "Bolivia": "Plurinational State of Bolivia", "Congo, Rep.": "Republic of the Congo", "Congo, Dem. Rep.": "Democratic Republic of the Congo", "Cote d'Ivoire" : "Cфte d'Ivoire", "Curacao": "Curaзao/Netherlands Antilles", "Korea, Dem. People's Rep.": "Democratic People's Republic of Korea",
                    "Egypt, Arab Rep.": "Egypt", "Hong Kong SAR, China": "Hong Kong (China)", "Iran, Islamic Rep.": "Islamic Republic of Iran", "Korea, Rep.": "Korea", "Kyrgyz Republic": "Kyrgyzstan", "Lao PDR": "Lao People's Democratic Republic", "Moldova": "Republic of Moldova", "North Macedonia": "Republic of North Macedonia", "Tanzania": "United Republic of Tanzania",
                    "Turkiye": "Turkey", "Venezuela, RB": "Bolivarian Republic of Venezuela", "Vietnam": "Viet Nam", "Yemen, Rep.": "Yemen"
                    }  # Словарь сравнительной таблицы