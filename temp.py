import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import re
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import os

path_to_file = 'data/БРИТ/БРИТ Госзадание №1 2022.xlsx'
name_spo = 'БРИТ'


df_1 = pd.read_excel(path_to_file)

# Создаем мультииндекс и
ml_1 = df_1[['Наименование государственной услуги','Категория потребителей государственной услуги',
                                       'Уникальный номер реестровой записи','Профессии и специальности по программам среднего профессионального образования']]

data_1 = df_1[['Категория потребителей','Формы образования и формы реализации образовательных программ','Госзадание №1']]

ml_df1 = pd.MultiIndex.from_frame(ml_1,names=['Наименование государственной услуги','Категория потребителей государственной услуги',
                                       'Уникальный номер реестровой записи','Профессии и специальности по программам среднего профессионального образования'])

data_1.index = ml_df1

df_2 = pd.read_excel('data/БРИТ/БРИТ Госзадание №2 2022.xlsx')

# Создаем мультииндекс и
ml_2 = df_2[['Наименование государственной услуги','Категория потребителей государственной услуги',
                                       'Уникальный номер реестровой записи','Профессии и специальности по программам среднего профессионального образования']]

data_2 = df_2[['Госзадание №2']]

ml_df2 = pd.MultiIndex.from_frame(ml_2,names=['Наименование государственной услуги','Категория потребителей государственной услуги',
                                       'Уникальный номер реестровой записи','Профессии и специальности по программам среднего профессионального образования'])

data_2.index = ml_df2

temp = pd.DataFrame()



gos_zad_df=pd.concat([data_1,data_2],axis=1,join='outer')

gos_report = pd.read_excel('data/БРИТ/БРИТ Госзадание Отчет 2022.xlsx')

# Создаем мультииндекс
ml_report_df = gos_report[['Наименование государственной услуги','Категория потребителей государственной услуги',
                                       'Уникальный номер реестровой записи','Профессии и специальности по программам среднего профессионального образования']]

data_report = gos_report[['исполнено на отчетную дату','причина отклонения']]

ml_report = pd.MultiIndex.from_frame(ml_report_df,names=['Наименование государственной услуги','Категория потребителей государственной услуги',
                                       'Уникальный номер реестровой записи','Профессии и специальности по программам среднего профессионального образования'])

data_report.index = ml_report

itog_df = pd.concat([gos_zad_df,data_report],axis=1,join='outer')

itog_df['Отклонение в абсолютных единицах'] = (itog_df['Госзадание №2']-itog_df['исполнено на отчетную дату']).abs()

itog_df['Отклонение в процентах'] = (itog_df['Отклонение в абсолютных единицах'] * 100 / itog_df['Госзадание №2']).round(2)

itog_df['Отклонение больше 5%'] =  itog_df['Отклонение в процентах'].apply(lambda x: 'Да'if x >5 else 'Нет')

itog_df['Отклонение в процентах'] = itog_df['Отклонение в процентах'].astype(str) + '%'

itog_df = itog_df.reset_index()

itog_df.insert(0,'Наименование ПОО',name_spo)

itog_df.to_excel('Итог.xlsx',index=False)



