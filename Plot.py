import openpyxl
from openpyxl.chart import Reference, BarChart, LineChart

file = r'C:\Users\Артем\Desktop\Industry consumption.xlsx'

df = openpyxl.load_workbook(file)
sheet = df['Total']

chart = LineChart()

chart.anchor="J5"
chart.width=15 # in cm
chart.height=5 # in cm

data = Reference(sheet, min_col=34, max_col=64, min_row=1, max_row=159)
chart.add_data(data)
chart.width = 30  # Ширина и высота графика (в см)
chart.height = 10

sheet.add_chart(chart, "E5")  # Добавляем график на лист в переменной sheet с левым верхним углом в ячейке Е5
chart.marker.symbol = "round"
df.save(file)