import pandas as pd
import os
import time
import shutil
from datetime import datetime

def education():
    # file_path_edu = os.sep * 2 + os.path.join("tg-storage01", "Аналитический отдел", "Личные", "Федорова",
    #                                           "Служба персонала", "Действующие сотрудники", "Образование",
    #                                           "Образование - выгрузка.xlsx")
    # file_path_res = os.sep * 2 + os.path.join("tg-storage01", "Аналитический отдел", "Личные", "Федорова",
    #                                           "Служба персонала", "Действующие сотрудники", "Образование",
    #                                           "Образование.xlsx")
    # file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Аналитический отдел", "Личные", "Федорова",
    #                                                 "Служба персонала", "Действующие сотрудники",
    #                                                 "Действующие сотрудники.xlsx")
    file_path_edu = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Отдел аналитики",
                                              "Выгрузки. Действующие сотрудники", "Образование",
                                              "Образования сотрудников - рассылка XLSX.xlsx")
    file_path_res = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Отдел аналитики",
                                              "Выгрузки. Действующие сотрудники", "Образование",
                                              "Образование.xlsx")
    file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Кадровый учет", "Действующие сотрудники.xlsx")

    # file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Отдел аналитики", "Действующие сотрудники.xlsx")

    df_edu = pd.read_excel(file_path_edu)
    df_edu = df_edu.sort_values(['Табельный номер (с префиксами)', 'Окончание'], ascending=(True, False))

    def combine_education(group):
        education = ''
        for index, row in group.iterrows():
            specialty = row["Специальность"]
            institution = row["Учебное заведение"]
            ending = row["Окончание"]
            type = row["Вид образования"]
            if pd.isna(specialty) and pd.isna(institution):
                return None
            if pd.isna(specialty):
                specialty = ' '
            else:
                if pd.isna(institution):
                    specialty = specialty
                else:
                    specialty = f' - {specialty}'
            if pd.isna(institution):
                institution = ''
            else:
                institution = institution
            if pd.isna(ending):
                ending = ''
            else:
                ending = f' ({pd.to_datetime(ending, format="%d.%m.%Y").year})'
            education += f'{institution}{specialty} ({type}){ending}; '
        if education:
            return education[:-2]
        else:
            return None
    result = df_edu.groupby(['Табельный номер (с префиксами)', 'Сотрудник']).apply(combine_education).reset_index(name='Образование')
    result = result.dropna()
    result.to_excel(file_path_res, index=False)

    # загрузка файлов для добавления инф об обр в осн файл
    df_emp = pd.read_excel(file_path_employees)
    df_edu = pd.read_excel(file_path_res)
    # df_emp.to_excel('file_path_employees.xlsx', index=False)
    # df_edu.to_excel('file_path_res.xlsx', index=False)
    # объединение таблиц по столбцу "Табельный номер"
    merged_df = pd.merge(df_emp, df_edu,
                         on='Табельный номер (с префиксами)', how='left')
    #
    # вставка столбца "Образование" перед первым столбцом
    # merged_df.to_excel('merged_df.xlsx', index=False)
    # merged_df = merged_df.drop(['Образование_x', 'Образование_y'], axis=1)

    # merged_df.insert(15, 'Образование', merged_df.pop('Образование'))

    # сохранение результата в файл
    merged_df.to_excel(file_path_employees, index=False)

# while True:
#     education()






# while True:
#     if datetime.now().hour == 16:
#         print(datetime.now())
#         education()
#         employees()
#         time.sleep(60 * 60 * 12)
#     else:
#         time.sleep(60 * 30)