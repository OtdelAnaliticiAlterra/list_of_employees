import os
import re
import pandas as pd

# Указываем путь до папки с файлами
# path_to_photos = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Кадровый учет",
#                                           "Дейст. сотрудники", "фото сотрудников")

def photos():

    # file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Аналитический отдел", "Личные", "Федорова",
    #                                                 "Служба персонала", "Действующие сотрудники",
    #                                                 "Действующие сотрудники.xlsx")
    path_to_photos = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Кадровый учет",
                                              "Дейст. сотрудники", "фото сотрудников", "ЦБ")
    file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Отдел аналитики", "Действующие сотрудники.xlsx")

    # Создаем пустой список для хранения результатов
    files_list = []

    # Обходим все папки и файлы внутри указанной директории
    for root, dirs, files in os.walk(path_to_photos):
        for file in files:
            # Получаем полный путь до файла
            file_path = os.path.join(root, file)
            # Получаем имя файла, удаляем лишние пробелы и сохраняем его в переменную
            file_name = file.split()[0].strip()
            # Ищем в имени файла последовательность из трех цифр после пробела
            if re.search(r'\d{3}$', file_name):
                # Добавляем имя файла и его полный путь в список
                files_list.append({'file_name': file_name, 'file_path': file_path})

    # Создаем DataFrame из списка файлов
    # print(files_list)
    df_photos = pd.DataFrame(files_list)
    # print(df_photos)
    df_emp = pd.read_excel(file_path_employees)
    # print(df_emp)

    # Удаляем дубликаты по file_name
    df_photos.drop_duplicates(subset=['file_name'], inplace=True)

    merged_df = pd.merge(df_emp, df_photos, left_on='Табельный номер (с префиксами)', right_on='file_name',
                         how='left')

    merged_df.drop('file_name', axis=1, inplace=True)


    # merged_df.to_excel(file_path_employees, index=False)
    # writer = pd.ExcelWriter(file_path_employees, engine='openpyxl')
    merged_df.to_excel(file_path_employees, index=False)

# photos()