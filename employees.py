import pandas as pd
import os
import time
import shutil
from datetime import datetime

def employees():
    # file_path_for_copy = os.sep * 2 + os.path.join("tg-storage01", "Аналитический отдел", "Личные", "Федорова",
    #                                               "Служба персонала", "Действующие сотрудники", "Список ЗУП",
    #                                               "ЗУП.xlsx")
    # file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Аналитический отдел", "Личные", "Федорова",
    #                                           "Служба персонала", "Действующие сотрудники")
    file_path_for_copy = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Отдел аналитики",
                                                   "Выгрузки. Действующие сотрудники", "Список ЗУП", "Действующие сотрудники XLSX.xlsx")
    file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Отдел аналитики")
    file_to_change = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Отдел аналитики",
                                               "Выгрузки. Действующие сотрудники", "Список ЗУП", "Дата приема.xlsx")
    file_name = "Действующие сотрудники.xlsx"
    file_path_employees = os.path.join(file_path_employees, file_name)


    shutil.copyfile(file_path_for_copy, file_path_employees)
    df = pd.read_excel(file_path_employees, header=0)
    groups = df.groupby(['ФИО'])

    def select_rows(group):
        if len(group) == 1:
            return group
        elif 'Основное место работы' in group['Вид занятости'].values:
            return group[group['Вид занятости'] == 'Основное место работы']
        else:
            return group
    df_filtered = groups.apply(select_rows).reset_index(drop=True)

    def update_hire_date(row):
        tab_number = row['Табельный номер (с префиксами)'] if 'Табельный номер (с префиксами)' in row else row[
            'Таб. номер']
        hire_date = row['Дата приема']

        if tab_number in df_to_change['Таб. номер'].values:
            new_hire_date = df_to_change.loc[df_to_change['Таб. номер'] == tab_number, 'Дата приема'].values[0]
            return new_hire_date
        else:
            return hire_date

    df_to_change = pd.read_excel(file_to_change, header=0)

    df_filtered['Дата приема'] = df_filtered.apply(update_hire_date, axis=1)
    # df_filtered = groups.apply(select_rows).reset_index(drop=True)
    df_filtered.to_excel(file_path_employees, index=False)



    # df_filtered.to_excel(file_path_employees, index=False)
