import xlrd
import csv
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

xls_file_path = r"C:\Users\USER\Desktop\practice\city-air-pullution-index\archive\air_poll.xls"

csv_file_path = r'C:\Users\USER\Desktop\practice\city-air-pullution-index\data\air_poll.csv'


workbook = xlrd.open_workbook(xls_file_path)


sheet = workbook.sheet_by_index(0)

header = sheet.row_values(2)

with open(csv_file_path, 'w', newline='',encoding='utf-8') as csv_file:
    csv_writer = csv.writer(csv_file)
    
    
    csv_writer.writerow(header)
    
    for row_idx in range(3, sheet.nrows):
        csv_writer.writerow(sheet.row_values(row_idx))


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
df_unpivot.rename(columns={"Year": "Year", "Region": "Region", "Total_emissions(тыс.тонн)": "Total_emissions"}, inplace=True)

print(df_unpivot)

data_types = df_unpivot.dtypes
print(data_types)

# Export to kazpop.csv
df_unpivot.to_csv(r'C:\Users\USER\Desktop\practice\city-air-pullution-index\data\air_poll_piv.csv', index=False)


import matplotlib.cm as cm

# Загрузка данных из CSV файла
file_path = r'C:\Users\USER\Desktop\practice\city-air-pullution-index\data\air_poll_piv.csv'
df = pd.read_csv(file_path)

# Сортировка DataFrame по убыванию значений 'Total_emissions'
df = df.sort_values(by='Total_emissions', ascending=False)

# Создание графика пирога
fig, ax = plt.subplots(figsize=(12, 8))

# Using 'tab20' colormap
cmap = cm.get_cmap('tab20', len(df))

wedges, texts, autotexts = ax.pie(df['Total_emissions'], labels=None, autopct='%1.1f%%', startangle=140, wedgeprops=dict(width=0.4), pctdistance=0.85, colors=cmap.colors)

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