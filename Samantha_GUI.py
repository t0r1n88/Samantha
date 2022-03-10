import pandas as pd
import  openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import os
from docxtpl import DocxTemplate
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import datetime
from datetime import date
from openpyxl.chart.label import DataLabelList
from openpyxl.chart import BarChart, Reference, PieChart, PieChart3D, Series
import warnings


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def select_file_data_calculate_budget_spo():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_data_calculate_budget_spo
    # Получаем путь к файлу
    name_file_data_calculate_budget_spo = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_end_folder_calculate_budget_spo():
    """
    Функция для выбора папки куда будет генерироваться файл
    :return:
    """
    global path_to_end_folder_calculate_budget_spo
    path_to_end_folder_calculate_budget_spo = filedialog.askdirectory()

def calculate_budget_spo():
    pass



if __name__=='__main__':
    window = Tk()
    window.title('ЦОПП Бурятия')
    window.geometry('700x860')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку подсчета бюджетных заявок
    tab_calculate_budget_spo = ttk.Frame(tab_control)
    tab_control.add(tab_calculate_budget_spo, text='Бюджетные заявки СПО')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Бюджетных заявок СПО
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_calculate_budget_spo,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nПодсчет бюджетных заявок СПО')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_calculate_budget_spo = resource_path('logo.png')
    img_calculate_budget_spo = PhotoImage(file=path_to_calculate_budget_spo)
    Label(tab_calculate_budget_spo,
          image=img_calculate_budget_spo
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_calculate_budget_spo = LabelFrame(tab_calculate_budget_spo, text='Подготовка')
    frame_data_for_calculate_budget_spo.grid(column=0, row=2, padx=10)


    # Создаем кнопку Выбрать файл с данными
    btn_data_doc = Button(frame_data_for_calculate_budget_spo, text='1) Выберите таблицу с данными', font=('Arial Bold', 20),
                          command=select_file_data_calculate_budget_spo
                          )
    btn_data_doc.grid(column=0, row=3, padx=10, pady=10)
    #
    # Создаем кнопку для выбора папки куда будет генерироваться файл

    btn_choose_end_folder_doc = Button(frame_data_for_calculate_budget_spo, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder_calculate_budget_spo
                                       )
    btn_choose_end_folder_doc.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку для создания документов из таблиц с произвольной структурой
    btn_create_files_other = Button(tab_calculate_budget_spo, text='Подсчитать',
                                    font=('Arial Bold', 20),
                                    command=calculate_budget_spo
                                    )
    btn_create_files_other.grid(column=0, row=5, padx=10, pady=10)
    window.mainloop()