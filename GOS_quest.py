from docx import Document
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import re






# Создаем регулярные выражения
name_table_re = re. compile(r'Наименование государственной услуги:.*?(Реализация\s.+?)[2]')
cat_table_re = re.compile(r'Категории потребителей государственной услуги:.+?Физические лица,(.+?)[3]')

doc = Document('БРИТ.docx')
# последовательность всех таблиц документа
all_tables = doc.tables
print('Всего таблиц в документе:', len(all_tables))

# создаем пустой словарь под данные таблиц
data_tables = {i:None for i in range(len(all_tables))}
# проходимся по таблицам

lst_df =[]
for i, table in enumerate(all_tables):

    # print('\nДанные таблицы №', i)
    # создаем список строк для таблицы `i` (пока пустые)
    data_tables[i] = [[] for _ in range(len(table.rows))]
    # проходимся по строкам таблицы `i`
    for j, row in enumerate(table.rows):
        # проходимся по ячейкам таблицы `i` и строки `j`
        for cell in row.cells:
            # добавляем значение ячейки в соответствующий
            # список, созданного словаря под данные таблиц
            data_tables[i][j].append(cell.text)

    # смотрим извлеченные данные
    # (по строкам) для таблицы `i`
    df = pd.DataFrame(data_tables[i])
    lst_df.append(df)
dct_data = dict()

# Извлекаем
for i in range(1,len(lst_df),3):
    print(i)
    # Получаем совокупный текст таблиц с названием и категорией студентов
    text = lst_df[i].sum().sum()
    # Удаляем символы переноса
    text = text.replace('\n','')

    # Ищем название таблицы
    match_name = re.search(name_table_re,text)
    name_table = match_name.group(1).strip()
    # print(f'Название таблицы {name_table}')
    # Ищем категорию таблицы
    match_cat = re.search(cat_table_re,text)
    name_cat = match_cat.group(1).strip()
    # print(f'Категория таблицы {name_cat}')
    # print('*********************')

    # dct_data[name_table] ={name_cat:''}
    if name_table not in dct_data:
        dct_data[name_table]={name_cat:''}
    else:
        dct_data[name_table].update({name_cat:''})
    if i == 1:
        # удаляем первые 4 строки в датафрейме
        df =  lst_df[3].drop(labels=[0,1,2,3],axis=0)
        # избавляемся от знаков переноса строки

        df = df.applymap(lambda x: x.replace('\n', ''))

    else:
        # удаляем первые 4 строки в датафрейме
        df = lst_df[i+2].drop(labels=[0,1,2,3],axis=0)
        # избавляемся от знаков переноса строки
        df = df.applymap(lambda x:x.replace('\n',''))
    #Добавляем датафрейм в словарь
    dct_data[name_table][name_cat] = df



# print(dct_data)

for key,value in dct_data.items():
    print(key)
    print('**********')
    print(value)
    print(len(value))