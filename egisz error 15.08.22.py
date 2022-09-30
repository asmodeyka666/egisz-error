from tkinter import filedialog as fd
from tkinter import ttk
import tkinter as tk
import pandas as pd
import os
import time
import numpy as np
from itertools import cycle, islice
from openpyxl import Workbook


global SNILS_MR
    # SNILS dictionary
SNILS_MR = {"ГБУЗ 'Больница ДЗМ'" : '000000000000'
}


def daily():
    file_names_zip = fd.askopenfilename(multiple=True)
    if os.path.exists(os.path.dirname(file_names_zip[0]) + '\Раздача') == False:
        os.mkdir(os.path.dirname(file_names_zip[0]) + '\Раздача')
    print("Все файлы для обработки:", file_names_zip, sep='/n')
    for file_name in file_names_zip:
        print("Рассматриваем: ", os.path.basename(file_name))
        if '2196' in file_name:
            modify(file_name)
        elif '2244' in file_name:
            modify_one(file_name)
        elif 'НеИД' in file_name or '2389' in file_name:
            modify_noID(file_name)
        
# Разбивка Идов
def modify(file_name):
    if file_name == '':
        file_name = fd.askopenfilename()

    # забираем файл
    xls = pd.ExcelFile(file_name)

    list_send_1 = ''
    list_send_2 = ''
    for lists_xls in xls.sheet_names:
        if '1' in lists_xls:
            if 'пере' in lists_xls or 'от' in lists_xls:
                list_send_1 = lists_xls
            else:
                list_1 = lists_xls
        elif '2' in lists_xls:
            if 'пере' in lists_xls or 'от' in lists_xls:
                list_send_2 = lists_xls
            else:
                list_2 = lists_xls
            
    df1 = pd.read_excel(file_name, list_1, dtype=str)
    df2 = pd.read_excel(file_name, list_2, dtype=str)
    if list_send_1 != '':
        df1_send_1 = pd.read_excel(file_name, list_send_1, dtype=str)
    if list_send_2 != '':
        df1_send_2 = pd.read_excel(file_name, list_send_2, dtype=str)
    df1 = pd.concat([df1, df2], ignore_index=True)
    df1.insert(0, 'Внесено', '')
    df1.insert(1, 'Комментарии', '')
    # не внесено
    df_not_included = df1.sort_values(['Описание ошибки', 'Наменование МО', 'Фамилия пациента', 'Имя пациента', 'Отчество', 'Дата рождения'],
                                           ascending=[True, True, True, True, True, True])
    df1_not_included = df_not_included.loc[(df_not_included['Кратность вакцинации'] == 'V1') | (df_not_included['Кратность вакцинации'] == 'переотправка V1')]
    df2_not_included = df_not_included.loc[(df_not_included['Кратность вакцинации'] == 'V2') | (df_not_included['Кратность вакцинации'] == 'переотправка V2')]        
    all_sheets = {list_1: df1_not_included, list_2 : df2_not_included}

    writer = pd.ExcelWriter(os.path.dirname(file_name) + r'\не внесено_' + os.path.basename(file_name),
                                engine='xlsxwriter')

    for sheet_name in all_sheets.keys():
        all_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

    writer.save()

    df1['СНИЛС МР верный'] = ''
    df1 = df1.fillna('')
    sortdf = ['Внесено', 'Комментарии', 'ИД пациента', 'Описание ошибки', 'Кратность вакцинации', 'СНИСЛ (из ЕРП)',
              'СНИЛС пациента из документа',
              'Фамилия пациента', 'Имя пациента', 'Отчество', 'Пол', 'Дата рождения', 'Наименование типа ДУЛ',
              'Серия ДУЛ',
              'Номер ДУЛ', 'Дата выдачи', 'Полис ОМС', 'Контакный телефон', 'Мобильный телефон',
              'Адрес_регистрации_город',
              'Адрес_регистрации_улица', 'Адрес_регистрации_дом', 'Адрес_проживания_город', 'Адрес_проживания_улица',
              'Адрес_проживания_дом', 'Дата вакцинации оцифровка', 'Наменование МО', 'СНИЛС МР',
              'Температура тела', 'Препарат вакцины', 'GTIN', 'Серийный номер (ISN)', 'Серия и контрольный номер',
              'Дата вакцинации', 'Статус передачи', 'Код типа ДУЛ', 'Наименование МУ', 
              'Допуск к вакцинации', 'Производитель', 'Срок годности', 'Жалобы на момент осмотра', 'Место введения',
              'Наличие противопоказаний', 'Побочная реакциия на прививку', 'ФИО МР', 'СНИЛС МР верный', 'DOCUMENT_ID', 'DOCUMENT_CREATED',
              'CCT', 'Адрес_регистрации_корпус', 'Адрес_регистрации_строение', 'Адрес_регистрации_квартира',
              'Адрес_регистрации_строкой', 'Адрес_регистрации_код_кладр', 'Адрес_проживания_корпус',
              'Адрес_проживания_строение', 'Адрес_проживания_квартира', 'Адрес_проживания_строкой',
              'Адрес_проживания_код_КЛАДР']

    #Заменяем СНИЛС мед работника на значение из базы
    df1['СНИЛС МР верный'] = df1['Наменование МО'].map(SNILS_MR)
    df1['СНИЛС МР верный'].fillna(df1['СНИЛС МР'], inplace=True)
    
    df1 = df1[sortdf]
    df1['GTIN'] = df1['GTIN'].str.replace('046', '')
    df1['Пол'] = df1['Пол'].str.replace('1', 'M')
    df1['Пол'] = df1['Пол'].str.replace('2', 'Ж')

    df1['Дата вакцинации оцифровка'] = df1['Дата вакцинации оцифровка'].combine(df1['Дата вакцинации'],
                                                                                (lambda x1, x2: x1 if x1 != '' else x2))
    
    df1_filtr = df1.sort_values(['Кратность вакцинации', 'Описание ошибки', 'Наменование МО', 'Фамилия пациента', 'Имя пациента', 'Отчество', 'Дата рождения'],
                                           ascending=[True, True, True, True, True, True, True])

    df1_filtr['Дата рождения'] = df1_filtr['Дата рождения'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
    df1_filtr['Дата выдачи'] = df1_filtr['Дата выдачи'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
    df1_filtr['Дата вакцинации оцифровка'] = df1_filtr['Дата вакцинации оцифровка'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
    

    df1_filtr['Дата рождения'] = (pd.to_datetime(df1_filtr['Дата рождения'], errors='coerce')
                          .dt.strftime("%d.%m.%Y"))
    
    df1_filtr['Дата выдачи'] = (pd.to_datetime(df1_filtr['Дата выдачи'], errors='coerce')
                          .dt.strftime("%d.%m.%Y"))
    
    df1_filtr['Дата вакцинации оцифровка'] = (pd.to_datetime(df1_filtr['Дата вакцинации оцифровка'], errors='coerce')
                          .dt.strftime("%d.%m.%Y"))
    
    
    df1_filtr = df1_filtr.iloc[:, :33]

    #Проверяем на различие СНИЛСов пациентов
    df1_filtr.insert(7, 'Сравнение СНИЛС', '')
    df1_filtr['Сравнение СНИЛС'] = ((df1_filtr.loc[:, 'СНИСЛ (из ЕРП)'] == df1_filtr.loc[:, 'СНИЛС пациента из документа']) |
                                    (df1_filtr['СНИСЛ (из ЕРП)'] == '') | (df1_filtr['СНИЛС пациента из документа'] == ''))
    
        
    df1_filtr.loc[df1_filtr['Кратность вакцинации'].isin(['V1', 'V2'])].to_excel(os.path.dirname(file_name) + '\Раздача' + r'\Раздача_' + os.path.basename(file_name), index=False)
    
    df1_filtr_not_included = df1_filtr.loc[df1_filtr['Кратность вакцинации'].isin(['переотправка V1', 'переотправка V2'])]
    df1_filtr_not_included = df1_filtr_not_included.sort_values(['Кратность вакцинации', 'Фамилия пациента', 'Имя пациента', 'Отчество', 'Дата рождения'],
                                           ascending=[True, True, True, True, True])
    if len(df1_filtr_not_included) > 0:
        df1_filtr_not_included.to_excel(os.path.dirname(file_name) + r'\Переотправка_' + os.path.basename(file_name), index=False)
    

    os.startfile(os.path.dirname(file_name))



def modify_one(file_name):
    if file_name == '':
        file_name = fd.askopenfilename()

    # забираем файл
    df1 = pd.read_excel(file_name, dtype=str)
    
    df1 = df1.fillna('')
    df1['СНИЛС МР верный'] = ''
    sortdf = ['ID_EMIAS', 'ERROR_DESCRIPTION', 'Кратность вакцинации', 'СНИСЛ (из ЕРП)',
              'СНИЛС пациента из документа',
              'Фамилия пациента', 'Имя пациента', 'Отчество', 'Пол', 'Дата рождения', 'Наименование типа ДУЛ',
              'Серия ДУЛ',
              'Номер ДУЛ', 'Дата выдачи', 'Полис ОМС', 'Контакный телефон', 'Мобильный телефон',
              'Адрес_регистрации_город',
              'Адрес_регистрации_улица', 'Адрес_регистрации_дом', 'Адрес_проживания_город', 'Адрес_проживания_улица',
              'Адрес_проживания_дом', 'Дата вакцинации оцифровка', 'Наменование МО', 'СНИЛС МР',
              'Температура тела', 'Препарат вакцины', 'GTIN', 'Серийный номер (ISN)', 'Серия и контрольный номер',
              'Дата вакцинации', 'STATUS_ERROR', 'Код типа ДУЛ', 'Наименование МУ', 
              'Допуск к вакцинации', 'Производитель', 'Срок годности', 'Жалобы на момент осмотра', 'Место введения',
              'Наличие противопоказаний', 'Побочная реакциия на прививку', 'ФИО МР', 'СНИЛС МР верный', 'DOCUMENT_ID', 'DOCUMENT_CREATED',
              'CCT', 'Адрес_регистрации_корпус', 'Адрес_регистрации_строение', 'Адрес_регистрации_квартира',
              'Адрес_регистрации_строкой', 'Адрес_регистрации_код_кладр', 'Адрес_проживания_корпус',
              'Адрес_проживания_строение', 'Адрес_проживания_квартира', 'Адрес_проживания_строкой',
              'Адрес_проживания_код_КЛАДР']

    #Заменяем СНИЛС мед работника на значение из базы
    df1['СНИЛС МР верный'] = df1['Наменование МО'].map(SNILS_MR)
    df1['СНИЛС МР верный'].fillna(df1['СНИЛС МР'], inplace=True)


    df1 = df1[sortdf]
    df1['GTIN'] = df1['GTIN'].str.replace('046', '')
    df1['Пол'] = df1['Пол'].str.replace('1', 'M')
    df1['Пол'] = df1['Пол'].str.replace('2', 'Ж')

    df1['Дата вакцинации оцифровка'] = df1['Дата вакцинации оцифровка'].combine(df1['Дата вакцинации'],
                                                                                (lambda x1, x2: x1 if x1 != '' else x2))
    
    df1_filtr = df1.sort_values(['Фамилия пациента', 'Имя пациента', 'Дата рождения'],
                                           ascending=[True, True, True])
    
    df1_filtr['Дата рождения'] = df1_filtr['Дата рождения'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
    df1_filtr['Дата выдачи'] = df1_filtr['Дата выдачи'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
    df1_filtr['Дата вакцинации оцифровка'] = df1_filtr['Дата вакцинации оцифровка'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
    
    df1_filtr['Дата рождения'] = (pd.to_datetime(df1_filtr['Дата рождения'], errors='coerce')
                          .dt.strftime("%d.%m.%Y"))
    df1_filtr['Дата выдачи'] = (pd.to_datetime(df1_filtr['Дата выдачи'], errors='coerce')
                          .dt.strftime("%d.%m.%Y"))
    df1_filtr['Дата вакцинации оцифровка'] = (pd.to_datetime(df1_filtr['Дата вакцинации оцифровка'], errors='coerce')
                          .dt.strftime("%d.%m.%Y"))
    
    df1_filtr = df1_filtr.iloc[:, :31]

    df1_filtr.to_excel(os.path.dirname(file_name) + '\Раздача' + r'\Пересобранный_' + os.path.basename(file_name), index=False)





def modify_noID(file_name):
    if file_name == '':
        file_name = fd.askopenfilename()

    if 'v1_по_v2' not in file_name and 'v1' not in file_name and 'v2' not in file_name and 'V1' not in file_name and 'V2' not in file_name and 'в1' not in file_name and 'В1' not in file_name and 'в2' not in file_name and 'В2' not in file_name:
    # забираем файл
        xls = pd.ExcelFile(file_name)

        list_1 = ''
        list_2 = ''
        for lists_xls in xls.sheet_names:
            if '1' in lists_xls:
                if 'пере' in lists_xls or 'от' in lists_xls:
                    list_send_1 = lists_xls
                else:
                    list_1 = lists_xls
            elif '2' in lists_xls:
                if 'пере' in lists_xls or 'от' in lists_xls:
                    list_send_2 = lists_xls
                else:
                    list_2 = lists_xls
            
        if list_1 != '':
            df1 = pd.read_excel(file_name, list_1, dtype=str)
        else:
            df1 = pd.DataFrame()
        if list_2 != '':
            df2 = pd.read_excel(file_name, list_2, dtype=str)
        else:
            df2 = pd.DataFrame()
        df1 = pd.concat([df1, df2], ignore_index=True)

        df1.insert(0, 'Внесено', '')
        df1.insert(1, 'Комментарии', '')
        # не внесено
        df_not_included = df1.sort_values(['Кратность вакцинации', 'Описание ошибки', 'Наменование МО', 'Фамилия', 'Имя', 'Отчество', 'Дата рождения'],
                                           ascending=[True, True, True, True, True, True, True])
        df1_not_included = df_not_included.loc[(df_not_included['Кратность вакцинации'] == 'V1') | (df_not_included['Кратность вакцинации'] == 'переотправка V1')]
        df2_not_included = df_not_included.loc[(df_not_included['Кратность вакцинации'] == 'V2') | (df_not_included['Кратность вакцинации'] == 'переотправка V2')]        
        all_sheets = {'V1' : df1_not_included, 'V2' : df2_not_included}

        writer = pd.ExcelWriter(os.path.dirname(file_name) + r'\не внесено_' + os.path.basename(file_name),
                                engine='xlsxwriter')

        for sheet_name in all_sheets.keys():
            all_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

        writer.save()

        df1['СНИЛС МР верный'] = ''
        df1 = df1.fillna('')
        sortdf = ['Внесено', 'Комментарии', 'UNIDENTIFIED_PATIENT_ID', 'Описание ошибки', 'Кратность вакцинации', 'СНИЛС',
              'Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'Наименование ДУЛ',
              'Серия дул', 'Номер ДУЛ', 'Дата выдачи ДУЛ', 'Наименование документа иностарнного гаржданина', 'Серия документа иностранного гражданина',
              'Номер документа иностранного гражданина', 'Полис ОМС', 'Город',
              'Улица', 'Дом', 'Дата вакцинации', 'Наменование МО', 'СНИЛС МР',
              'Препарат вакцины', 'GTIN', 'Серийный номер (ISN)', 'Серия и контрольный номер',
              'DOCUMENT_CREATED', 'CCT', 'Признак иностранного гражданина', 'Код типа ДУЛ', 'Код типа документа иностранного гражданина',
              'Регион', 'Наименование МУ', 'Допуск к вакцинации', 'ФИО МР', 'СНИЛС МР верный', 'Дата вакцинации', 'Производитель']
        #Заменяем СНИЛС мед работника на значение из базы
        df1['СНИЛС МР верный'] = df1['Наменование МО'].map(SNILS_MR)
        df1['СНИЛС МР верный'].fillna(df1['СНИЛС МР'], inplace=True)

        df1 = df1[sortdf]
        df1['GTIN'] = df1['GTIN'].str.replace('046', '')

        df1_filtr = df1.sort_values(['Кратность вакцинации', 'Описание ошибки', 'Наменование МО', 'Фамилия', 'Имя', 'Отчество', 'Дата рождения'],
                                           ascending=[True, True, True, True, True, True, True])
        
        df1_filtr = df1_filtr.iloc[:, :-12]
        
        df1_filtr['Дата рождения'] = df1_filtr['Дата рождения'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
        df1_filtr['Дата выдачи ДУЛ'] = df1_filtr['Дата выдачи ДУЛ'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
        df1_filtr['Дата вакцинации'] = df1_filtr['Дата вакцинации'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
        

        df1_filtr['Дата рождения'] = (pd.to_datetime(df1_filtr['Дата рождения'], errors='coerce')
                          .dt.strftime("%d.%m.%Y"))
        
        df1_filtr['Дата выдачи ДУЛ'] = (pd.to_datetime(df1_filtr['Дата выдачи ДУЛ'], errors='coerce')
                          .dt.strftime("%d.%m.%Y"))
        
        df1_filtr['Дата вакцинации'] = (pd.to_datetime(df1_filtr['Дата вакцинации'], errors='coerce')
                          .dt.strftime("%d.%m.%Y"))

        df1_filtr.loc[df1_filtr['Кратность вакцинации'].isin(['V1', 'V2'])].to_excel(os.path.dirname(file_name) + '\Раздача' + r'\Раздача_' + os.path.basename(file_name), index=False)
        df1_filtr_not_included = df1_filtr.loc[df1_filtr['Кратность вакцинации'].isin(['переотправка V1', 'переотправка V2'])]
        df1_filtr_not_included = df1_filtr_not_included.sort_values(['Кратность вакцинации', 'Фамилия', 'Имя', 'Отчество', 'Дата рождения'],
                                           ascending=[True, True, True, True, True])
        if len(df1_filtr_not_included) > 0:
            df1_filtr_not_included.to_excel(os.path.dirname(file_name) + r'\Переотправка_' + os.path.basename(file_name), index=False)
              

        root.destroy()

    else:
        df1 = pd.read_excel(file_name, dtype=str)
        df1 = df1.fillna('')
        df1['СНИЛС МР верный'] = ''
        sortdf = ['UNIDENTIFIED_PATIENT_ID', 'Описание ошибки', 'Кратность вакцинации', 'СНИЛС',
              'Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'Наименование ДУЛ',
              'Серия дул', 'Номер ДУЛ', 'Дата выдачи ДУЛ', 'Наименование документа иностарнного гаржданина', 'Серия документа иностранного гражданина',
              'Номер документа иностранного гражданина', 'Полис ОМС', 'Город',
              'Улица', 'Дом', 'Дата вакцинации', 'Наменование МО', 'СНИЛС МР',
              'Препарат вакцины', 'GTIN', 'Серийный номер (ISN)', 'Серия и контрольный номер',
              'DOCUMENT_CREATED', 'CCT', 'Признак иностранного гражданина', 'Код типа ДУЛ', 'Код типа документа иностранного гражданина',
              'Регион', 'Наименование МУ', 'Допуск к вакцинации', 'ФИО МР', 'СНИЛС МР верный', 'Дата вакцинации', 'Производитель']
        #Заменяем СНИЛС мед работника на значение из базы
        df1['СНИЛС МР верный'] = df1['Наменование МО'].map(SNILS_MR)
        df1['СНИЛС МР верный'].fillna(df1['СНИЛС МР'], inplace=True)

        df1 = df1[sortdf]
        df1['GTIN'] = df1['GTIN'].str.replace('046', '')

        df1_filtr = df1.sort_values(['Фамилия', 'Имя', 'Дата рождения'],
                                           ascending=[True, True, True])

        df1_filtr = df1_filtr.iloc[:, :26]
        
        df1_filtr['Дата рождения'] = df1_filtr['Дата рождения'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
        df1_filtr['Дата рождения'] = (pd.to_datetime(df1_filtr['Дата рождения'], errors='coerce', format = '%Y-%m-%d')
                          .dt.strftime("%d.%m.%Y"))
        df1_filtr['Дата выдачи ДУЛ'] = df1_filtr['Дата выдачи ДУЛ'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
        df1_filtr['Дата выдачи ДУЛ'] = (pd.to_datetime(df1_filtr['Дата выдачи ДУЛ'], errors='coerce', format = '%Y-%m-%d')
                          .dt.strftime("%d.%m.%Y"))
        df1_filtr['Дата вакцинации'] = df1_filtr['Дата вакцинации'].map(lambda x: x[0: x.find('T')] if 'T' in x else x)
        df1_filtr['Дата вакцинации'] = (pd.to_datetime(df1_filtr['Дата вакцинации'], errors='coerce', format = '%Y-%m-%d')
                          .dt.strftime("%d.%m.%Y"))

        df1_filtr.to_excel(os.path.dirname(file_name) + '\Раздача' + r'\Пересобранный_' + os.path.basename(file_name), index=False)




    
root = tk.Tk()
root.title("Обработка запросов")
root.geometry('600x400')
root["bg"] = "#fff"


button1 = tk.Button(text="Ежедневная модификация",
                   command=daily, background="#fff", foreground="#3b3e41",
                   padx="30", pady="15", font="15")
button1.pack()


root.mainloop()
