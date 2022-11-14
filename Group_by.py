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

# Задаем стартовые параметры для парсинга и фильтрации (начальный и конечный года, энергоресурсы, отрасль и путь к файлу)
file = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'
xl = pd.ExcelFile(file)  # Загружаем spreadsheet (электронную таблицу) в объект pandas
# print(xl.sheet_names) # Печатаем названия листов в данном файле
dfs = {

}

dfs = xl.parse(sheet_name='Total', skiprows=0, usecols='AH:BL')  # Парсим лист эксель-файла

xl.close()  # Закрываем читаемый файл

if os.path.exists(file_to_parse):
    mode = "a"
    if_sheet_exists = "overlay"
else:
    mode = "w"
    if_sheet_exists = None
writer = pd.ExcelWriter(file_to_parse, engine='openpyxl', mode=mode,
                        if_sheet_exists=if_sheet_exists)  # Указываем writer библиотеки
country_list = ['World', 'OECD Americas', 'OECD Asia Oceania', 'OECD Europe', 'Africa', 'Non-OECD Americas', 'Middle East', 'Non-OECD Europe and Eurasia', 'Non-OECD Asia (excluding China)', 'China (P.R. of China and Hong Kong, China)', 'World marine bunkers', 'World aviation bunkers', 'Albania', 'Algeria', 'Angola', 'Argentina', 'Armenia', 'Australia', 'Austria', 'Azerbaijan', 'Bahrain', 'Bangladesh', 'Belarus', 'Belgium', 'Benin', 'Plurinational State of Bolivia', 'Bosnia and Herzegovina', 'Botswana', 'Brazil', 'Brunei Darussalam', 'Bulgaria', 'Cambodia', 'Cameroon', 'Canada', 'Chile', "People's Republic of China", 'Colombia', 'Republic of the Congo', 'Costa Rica', "Cфte d'Ivoire", 'Croatia', 'Cuba', 'Curaзao/Netherlands Antilles', 'Cyprus', 'Czech Republic', "Democratic People's Republic of Korea", 'Democratic Republic of the Congo', 'Denmark', 'Dominican Republic', 'Ecuador', 'Egypt', 'El Salvador', 'Equatorial Guinea', 'Eritrea', 'Estonia', 'Ethiopia', 'Finland', 'France', 'Gabon', 'Georgia', 'Germany', 'Ghana', 'Gibraltar', 'Greece', 'Guatemala', 'Guyana', 'Haiti', 'Honduras', 'Hong Kong (China)', 'Hungary', 'Iceland', 'India', 'Indonesia', 'Islamic Republic of Iran', 'Iraq', 'Ireland', 'Israel', 'Italy', 'Jamaica', 'Japan', 'Jordan', 'Kazakhstan', 'Kenya', 'Korea', 'Kosovo', 'Kuwait', 'Kyrgyzstan', "Lao People's Democratic Republic", 'Latvia', 'Lebanon', 'Libya', 'Lithuania', 'Luxembourg', 'Malaysia', 'Malta', 'Mauritius', 'Mexico', 'Republic of Moldova', 'Mongolia', 'Montenegro', 'Morocco', 'Mozambique', 'Myanmar', 'Namibia', 'Nepal', 'Netherlands', 'New Zealand', 'Nicaragua', 'Niger', 'Nigeria', 'Republic of North Macedonia', 'Norway', 'Oman', 'Pakistan', 'Panama', 'Paraguay', 'Peru', 'Philippines', 'Poland', 'Portugal', 'Qatar', 'Romania', 'Russian Federation', 'Saudi Arabia', 'Senegal', 'Serbia', 'Singapore', 'Slovak Republic', 'Slovenia', 'South Africa', 'South Sudan', 'Spain', 'Sri Lanka', 'Sudan', 'Suriname', 'Sweden', 'Switzerland', 'Syrian Arab Republic', 'Chinese Taipei', 'Tajikistan', 'United Republic of Tanzania', 'Thailand', 'Togo', 'Trinidad and Tobago', 'Tunisia', 'Turkey', 'Turkmenistan', 'Ukraine', 'United Arab Emirates', 'United Kingdom', 'United States', 'Uruguay', 'Uzbekistan', 'Bolivarian Republic of Venezuela', 'Viet Nam', 'Yemen', 'Zambia', 'Zimbabwe']
group_list = [nan, nan, nan, nan, nan, nan, nan, nan, nan, nan, nan, nan, 4.0, 2.0, 5.0, 5.0, 6.0, 3.0, 3.0, 6.0, 5.0, 1.0, 4.0, 3.0, 2.0, 5.0, 5.0, 1.0, 5.0, 1.0, 4.0, 2.0, 5.0, 4.0, 1.0, 4.0, 3.0, 2.0, 5.0, 4.0, 3.0, 4.0, nan, 4.0, 4.0, nan, 4.0, 4.0, 3.0, 5.0, 3.0, 1.0, 4.0, nan, 4.0, 4.0, 3.0, 3.0, 2.0, 2.0, 3.0, 4.0, nan, 5.0, 1.0, 4.0, 2.0, 3.0, 4.0, 4.0, 3.0, 3.0, 5.0, 2.0, 4.0, 4.0, 5.0, 5.0, 5.0, 3.0, 4.0, 4.0, 5.0, 4.0, 3.0, 5.0, 4.0, nan, 4.0, 2.0, nan, 4.0, 4.0, 5.0, 4.0, 3.0, 5.0, 3.0, 4.0, 4.0, 3.0, 4.0, 4.0, 5.0, 2.0, 3.0, 4.0, 4.0, 2.0, 1.0, 4.0, 3.0, 2.0, 5.0, 4.0, 1.0, 5.0, 3.0, 4.0, 5.0, 4.0, 4.0, 6.0, 2.0, 4.0, 4.0, 2.0, 4.0, 3.0, 5.0, 2.0, 5.0, 4.0, 5.0, 4.0, 4.0, 3.0, 1.0, nan, 6.0, 3.0, nan, 4.0, 3.0, 5.0, 4.0, nan, 6.0, 5.0, 3.0, 4.0, 1.0, 6.0, nan, nan, 2.0, 4.0, 4.0]
cn = pd.DataFrame(list(zip(country_list, group_list)), columns=['Страна', 'Группа'])  # Создаем датафрейм из списков
#print(cn)
df = dfs  # Получаем лист из словаря dfs

table = df.loc[:, 'COUNTRY':'2019']
table = pd.merge(table, cn, left_on=['COUNTRY'], right_on=['Страна'],
                  how='left')
table.drop(['Страна'], axis='columns', inplace=True)
table1 = table['Группа']
# print(df1.name)  # Печатаем названия первичных ключей (названия столбцов) в данном массиве (не в датафрейме)
# print(df.keys())  # Печатаем названия первичных ключей (названия столбцов) в данном датафрейме
table1.to_excel(writer, sheet_name='Total', index=False, startcol=64)
# index=False отключает запись индексов, startcol=1 начианет запись с 1 стобца (нумерация с нуля).
writer.save()  # Сохраняем результат