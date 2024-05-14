import xlrd
import csv
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

xls_file_path = r"C:\Users\USER\Desktop\practice\city-air-pullution-index\archive\air_poll.xls"
csv_file_path = r'C:\Users\USER\Desktop\practice\city-air-pullution-index\data\air_poll.csv'

# Словарь с соответствиями названий регионов
region_translation = {
    'Абай': 'Abai Region',
    'Акмолинская': 'Akmola Region',
    'Актюбинская': 'Aktobe Region',
    'Алматинская': 'Almaty Region',
    'Атырауская': 'Atyrau Region',
    'Западно-Казахстанская': 'West Kazakhstan Region',
    'Жамбылская': 'Jambyl Region',
    'Жетісу': 'Jetisu Region',
    'Карагандинская': 'Karaganda Region',
    'Костанайская': 'Kostanay Region',
    'Кызылординская': 'Kyzylorda Region',
    'Мангистауская': 'Mangystau Region',
    'Павлодарская': 'Pavlodar Region',
    'Северо-Казахстанская': 'North Kazakhstan Region',
    'Туркестанская': 'Turkistan Region',
    'Ұлытау': 'Ulytau Region',
    'Восточно-Казахстанская': 'East Kazakhstan Region',
    'г. Астана': 'Astana city',
    'г. Алматы': 'Almaty city',
    'г.Шымкент': 'Shymkent city'
}

# Открытие файла Excel
workbook = xlrd.open_workbook(xls_file_path)
sheet = workbook.sheet_by_index(0)

# Получение заголовка
header = sheet.row_values(2)

# Открытие CSV файла для записи
with open(csv_file_path, 'w', newline='', encoding='utf-8') as csv_file:
    csv_writer = csv.writer(csv_file)
    
    # Запись заголовка
    csv_writer.writerow(header)
    
    # Запись данных
    for row_idx in range(3, sheet.nrows):
        row = sheet.row_values(row_idx)
        # Замена названий регионов на латиницу (предполагая, что регион находится в первой колонке)
        region_name = row[0]
        row[0] = region_translation.get(region_name, region_name)
        csv_writer.writerow(row)

csv = r"C:\Users\USER\Desktop\practice\city-air-pullution-index\data\air_poll.csv"
df = pd.read_csv(csv)
df.columns = df.columns.astype(str)
pd.set_option('display.max_columns', None)

# Unpivot the Excel data for CSV export
df_unpivot = pd.melt(df, id_vars='Unnamed: 0', value_vars=['2022.0'])
df_unpivot.rename(columns={"Unnamed: 0": "Region", "variable": "Year", "value": "Total_emissions(тыс.тонн)"}, inplace=True)
df_unpivot.replace("-", np.nan, inplace=True)
df_unpivot.dropna(axis=0, inplace=True)

# Exclude the row with region 'Республика Казахстан'
df_unpivot = df_unpivot[df_unpivot['Region'] != 'Республика Казахстан']

# Remove non-numeric characters from 'Year' column and then convert to integer
df_unpivot['Year'] = pd.to_numeric(df_unpivot['Year'].astype(str).str.replace(r'\D', ''), errors='coerce')

df_unpivot = df_unpivot.dropna(subset=['Year']).astype({'Year': int})

# Rename columns
df_unpivot.rename(columns={"Year": "Year", "Region": "Region", "Total_emissions(тыс.тонн)": "Value"}, inplace=True)

# Удаление столбца Year
df_unpivot = df_unpivot.drop(columns=['Year'])

print(df_unpivot)

data_types = df_unpivot.dtypes
print(data_types)


df_unpivot.to_csv(r'C:\Users\USER\Desktop\practice\city-air-pullution-index\data\air_poll_piv.csv', index=False)


import matplotlib.cm as cm

# Загрузка данных из CSV файла
file_path = r'C:\Users\USER\Desktop\practice\city-air-pullution-index\data\air_poll_piv.csv'
df = pd.read_csv(file_path)

# Сортировка DataFrame по убыванию значений 'Total_emissions'
df = df.sort_values(by='Value', ascending=False)

# Создание графика пирога
fig, ax = plt.subplots(figsize=(12, 8))

# Using 'tab20' colormap
cmap = cm.get_cmap('tab20', len(df))

wedges, texts, autotexts = ax.pie(df['Value'], labels=None, autopct='%1.1f%%', startangle=140, wedgeprops=dict(width=0.4), pctdistance=0.85, colors=cmap.colors)

# Установка местоположения текста внутри или снаружи пирога
ax.axis('equal')  # Убедитесь, что круг выглядит как круг
plt.title('Распределение выбросов загрязнений по регионам в 2022 году')

# Добавление легенды
legend = ax.legend(df['Region'], title='Области', bbox_to_anchor=(1, 0.5), loc="center right", bbox_transform=plt.gcf().transFigure)

# Изменение местоположения текста
for text, autotext in zip(texts, autotexts):
    text.set(size=10)  # размер текста
    autotext.set(size=8, color='white')  # размер и цвет текста с процентами
    autotext.set_horizontalalignment('center')  # выравнивание по центру

plt.tight_layout(rect=[0, 0, 0.85, 1])  # улучшение распределения местоположения графика в окне
plt.show()