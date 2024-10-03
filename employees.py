import pandas as pd
import os
import time
import shutil
from datetime import datetime

"""
    1. **Определение путей к файлам**:
       - Задаются пути к исходному файлу, который будет скопирован (`file_path_for_copy`), и к основному файлу с данными о сотрудниках (`file_path_employees`).
       - Определяется путь к файлу, содержащему новые даты приема сотрудников (`file_to_change`), и формируется окончательный путь к файлу с данными о сотрудниках.

    2. **Копирование файла**:
       - Исходный файл копируется в целевой путь, создавая новый файл с данными о действующих сотрудниках. 
       - Используется функция `shutil.copyfile` для выполнения копирования.

    3. **Загрузка данных**:
       - Загружается содержимое скопированного файла Excel в DataFrame.
       - Данные группируются по столбцу 'ФИО' для последующей обработки.

    4. **Фильтрация записей**:
       - Вспомогательная функция `select_rows` применяется к каждой группе сотрудников, чтобы оставить только одну запись для каждого сотрудника.
       - Если у сотрудника несколько записей, выбирается запись с 'Основным местом работы', если таковая имеется; в противном случае сохраняются все записи.

    5. **Обновление даты приема**:
       - Вспомогательная функция `update_hire_date` обновляет дату приема для сотрудников, основываясь на данных из файла с новыми датами приема.
       - Если табельный номер сотрудника присутствует в новом файле, дата приема обновляется на новую, иначе сохраняется старая дата.

    6. **Загрузка новых дат приема**:
       - Загружается файл с новыми датами приема в DataFrame (`df_to_change`), который используется для обновления информации.

    7. **Сохранение данных**:
       - Обновленные данные о сотрудниках сохраняются обратно в Excel-файл по пути `file_path_employees`.
       - Результирующий файл содержит актуальную информацию о действующих сотрудниках и их датах приема.

    8. **Результат**:
       - Функция завершает свою работу, обновляя файл с данными о сотрудниках.
"""
def employees():
    file_path_for_copy = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Отдел аналитики",
                                                   "Выгрузки. Действующие сотрудники", "Список ЗУП", "Действующие сотрудники XLSX.xlsx")
    file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Кадровый учет")

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
    df_filtered.to_excel(file_path_employees, index=False)



