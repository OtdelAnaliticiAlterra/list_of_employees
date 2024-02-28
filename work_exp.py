import pandas as pd
import os

def work_exp():
    # file_path_work = os.sep * 2 + os.path.join("tg-storage01", "Аналитический отдел", "Личные", "Федорова",
    #                                               "Служба персонала", "Действующие сотрудники", "Предыдущие места работы",
    #                                               "Предыдущие места работы.xlsx")
    file_path_work = os.sep * 2 + os.path.join("tg-storage01","Служба персонала", "Общие", "Отдел аналитики",
                                               "Выгрузки. Действующие сотрудники", "Предыдущие места работы", "Предыдущие места работы.xlsx")
    # file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Аналитический отдел", "Личные", "Федорова",
    #                                               "Служба персонала", "Действующие сотрудники", "Действующие сотрудники.xlsx")
    file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Кадровый учет", "Действующие сотрудники.xlsx")

    # file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Отдел аналитики", "Действующие сотрудники.xlsx")

    # загрузка файлов
    df_emp = pd.read_excel(file_path_employees)
    df_work = pd.read_excel(file_path_work)
    # удаление повторяющихся строк
    df_work = df_work.drop_duplicates()
    # объединение данных по ключу 1
    merged_df = pd.merge(df_emp, df_work, left_on='Табельный номер (с префиксами)', right_on='Табельный номер', how='left')
    merged_df['Сотрудник'] = merged_df['ФИО_x'].copy()
    merged_df['ФИО_x'] = merged_df['ФИО_x'].str.lower()
    df_work['ФИО'] = df_work['ФИО'].str.lower()
    # объединение данных по ключу 2
    merged_df = pd.merge(merged_df, df_work, left_on='ФИО_x', right_on='ФИО', how='left')
    merged_df['ФИО_x'] = merged_df['Сотрудник']
    # создание столбца "Работа"
    merged_df['Работа'] = merged_df.apply(lambda x: x['Предыдущие места работы (в обратной последовательности)_y'] if x['Предыдущие места работы (в обратной последовательности)_y'] else x['Предыдущие места работы (в обратной последовательности)_x'], axis=1)
    # удаление ненужных столбцов
    merged_df = merged_df.drop(['ФИО_y', 'Табельный номер_x', 'Сотрудник', 'ФИО', 'Табельный номер_y', 'Предыдущие места работы (в обратной последовательности)_x', 'Предыдущие места работы (в обратной последовательности)_y'], axis=1)
    merged_df = merged_df.rename(columns={'ФИО_x': 'ФИО'})
    # сохранение данных в файл
    merged_df.to_excel(file_path_employees, index=False)

