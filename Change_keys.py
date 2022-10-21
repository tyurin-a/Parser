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
country_dict = {"China" : "People's Republic of China", "Bolivia" : "Plurinational State of Bolivia", "Congo, Rep."
: "Republic of the Congo", "Congo, Dem. Rep." : "Democratic Republic of the Congo", "Cote d'Ivoire"
 : "Cфte d'Ivoire", "Curacao" : "Curaзao/Netherlands Antilles", "Korea, Dem. People's Rep." : "Democratic People's Republic of Korea",
"Egypt, Arab Rep." : "Egypt", "Hong Kong SAR, China" : "Hong Kong (China)", "Iran, Islamic Rep." : "Islamic Republic of Iran",
"Korea, Rep." : "Korea", "Kyrgyz Republic" : "Kyrgyzstan", "Lao PDR" : "Lao People's Democratic Republic",
"Moldova" : "Republic of Moldova", "North Macedonia" : "Republic of North Macedonia", "Tanzania" : "United Republic of Tanzania",
"Turkiye" : "Turkey", "Venezuela, RB" : "Bolivarian Republic of Venezuela", "Vietnam" : "Viet Nam", "Yemen, Rep." : "Yemen"
} # Словарь сравнительной таблицы

cn = pd.DataFrame(list(country_dict.items()), columns=['TableWB', 'TableIEA'])  # Создаем датафрейм из словаря
for k in dfs:
    df = dfs[k]  # Получаем лист из словаря dfs
    table2 = df.loc[:, 'Country Name':'Value added']
    table2 = pd.merge(table2, cn, left_on=['Country Name'], right_on=['TableWB'], how='left')  # Аналог ВПР для таблицы 2,
    # в которой будем менять ключи для адекватного соединения с 1ой таблицей с помощью сравнительной таблицы cn
    table2.drop(['TableWB'], axis='columns', inplace=True)
    table2.loc[pd.notna(table2['TableIEA']) == True, 'Country Name'] = table2['TableIEA']  # Меняем ключи в столбце 'Country Name'
    # таблицы 2 на значения столбца 'TableIEA', там где значение в строках столбца 'TableIEA' не пусто
    table2.drop(['TableIEA'], axis='columns', inplace=True)  # Удаляем лишний столбец
    # print(df1.name)  # Печатаем названия первичных ключей (названия столбцов) в данном массиве (не в датафрейме)
    # print(df.keys())  # Печатаем названия первичных ключей (названия столбцов) в данном датафрейме
    table2.to_excel(writer, sheet_name=str(k), index=False, startcol=10)
# index=False отключает запись индексов, startcol=1 начианет запись с 1 стобца (нумерация с нуля).
writer.save()  # Сохраняем результат