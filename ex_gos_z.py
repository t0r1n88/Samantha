from docx import Document
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import re
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import os

path = 'data/'
# перебираем папки
for dir in os.listdir(path):
    if os.path.isdir(f'{path}{dir}'):
        print(dir)
        #Перебираем файлы внутри папки Проверяем расширение файла и что он не временный
        for file in os.listdir(f'{path}{dir}'):
            print(file)


            if not file.startswith('~') and file.endswith('xlsx'):
                # получаем путь к файлу
                path_to_file = f'{path}{dir}/{file}'
                # получаем название файла

                name_file = file.split('.')[0]
                # Открываем документ
                df = pd.read_excel(path_to_file)
                print(df.head())





