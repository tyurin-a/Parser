# Импорт библиотек
import pandas as pd
import openpyxl
from openpyxl import load_workbook

def value_added():
    # Задаем стартовые параметры для парсинга и фильтрации (начальный и конечный года, энергоресурсы, отрасль и путь к файлу)
    start_year = '1990'
    end_year = '2019'
    file = r'C:\Users\Артем\Desktop\МФТИ\Магистратура\Диплом_Магистр\Industry (include construction) value added WB.xlsx'
    file_to_parse = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'

    xl = pd.ExcelFile(file)  # Загружаем spreadsheet (электронную таблицу) в объект pandas
    # print(xl.sheet_names) # Печатаем названия листов в данном файле

    # Загружаем лист (sheet_name) в DataFrame 'df', пропуская (по желанию) строки (skiprows) или столбцы (skipcols)
    df = xl.parse(sheet_name='Data', skiprows=3)
    # print(df.keys()) # Печатаем названия первичных ключей (названия столбцов) в данном датафрейме

    # Запись данных
    # Фильтрация данных
    country = (df['Country Name'])  # Отбор нужных столбцов

    writer = pd.ExcelWriter(file_to_parse, mode="a", engine='openpyxl',
                            if_sheet_exists='overlay')  # Указываем writer библиотеки

    for k in range(int(start_year), int(end_year) + 1, 1):
        data = (df[str(k)])  # Отбор нужных столбцов
        data.name = 'Value added'  # Переименовываем столбец (rename не работает, т.к. здесь он всего один)
        # # Запись в новый Excel-файл
        # book = load_workbook(file_to_parse)  # Получаем доступ к файлу MS Excel в который будем записывать датафрейм
        # writer.book = book  # Сохраняем предыдущую информацию файла, чтобы при записи она осталась
        # writer.sheets = dict((ws.title, ws) for ws in book.worksheets)  # ExcelWriter использует эту переменную для доступа к листу.
        # # Если оставить ее пустой, он не будет знать, что лист уже существует, и создаст новый лист.

        country.to_excel(writer, sheet_name=str(k), index=False, startcol=10)  # Записываем датафрейм в файл.
        # sheet_name='2019' показывает, в какой лист записываем.
        data.to_excel(writer, sheet_name=str(k), index=False, startcol=11)  # Записываем датафрейм в файл.
        # index=False отключает запись индексов, startcol=1 начианет запись с 1 стобца (нумерация с нуля).
    writer.save()  # Сохраняем результат

value_added()