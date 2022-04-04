from docx import Document
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import re
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import os

path_to_files = 'data/'
# перебираем папки
for dir in os.listdir(path_to_files):
    if os.path.isdir(f'{path_to_files}{dir}'):
        print(dir)
        #Перебираем файлы внутри папки Проверяем расширение файла и что он не временный
        for file in os.listdir(f'{path_to_files}{dir}'):
            print(file)

            if not file.startswith('~') and file.endswith('docx'):
                # получаем имя файла
                name_file = file.split('.')[0]
                # Открываем документ
                doc = Document(f'{path_to_files}{dir}/{file}')
                # последовательность всех таблиц документа
                all_tables = doc.tables
                print(f'Всего таблиц в документе {dir}:', len(all_tables))
                # # создаем пустой словарь под данные таблиц
                data_tables = {i:None for i in range(len(all_tables))}
                # проходимся по таблицам
                # Создаем список для храния полученных датафреймлв
                lst_df =[]
                for i, table in enumerate(all_tables):

                    print('\nДанные таблицы №', i)
                    # создаем список строк для таблицы `i` (пока пустые)
                    data_tables[i] = [[] for _ in range(len(table.rows))]
                    # проходимся по строкам таблицы `i`
                    for j, row in enumerate(table.rows):
                        # проходимся по ячейкам таблицы `i` и строки `j`
                        for cell in row.cells:
                            d = (cell.text)
                            # добавляем значение ячейки в соответствующий
                            # список, созданного словаря под данные таблиц
                            data_tables[i][j].append(cell.text)
                    temp_df = pd.DataFrame(data_tables[i])
                    lst_df.append(temp_df)
                    temp_df.to_excel(f'{name_file} {i}.xlsx')




# # doc = Document('БМК.docx')
# doc = Document('БКН.docx')
# # doc = Document('2.docx')
# d = 'Второе'
# # последовательность всех таблиц документа
# all_tables = doc.tables
# print('Всего таблиц в документе:', len(all_tables))
#
# # создаем пустой словарь под данные таблиц
# data_tables = {i:None for i in range(len(all_tables))}
# # проходимся по таблицам
#
# lst_df =[]
# for i, table in enumerate(all_tables):
#
#     print('\nДанные таблицы №', i)
#     # создаем список строк для таблицы `i` (пока пустые)
#     data_tables[i] = [[] for _ in range(len(table.rows))]
#     # проходимся по строкам таблицы `i`
#     for j, row in enumerate(table.rows):
#         # проходимся по ячейкам таблицы `i` и строки `j`
#         for cell in row.cells:
#             # добавляем значение ячейки в соответствующий
#             # список, созданного словаря под данные таблиц
#             data_tables[i][j].append(cell.text)
#     c = pd.DataFrame(data_tables[i])
#     c.to_excel(f'{d} {i}.xlsx')
#     # c.to_excel(f'{i}.xlsx')