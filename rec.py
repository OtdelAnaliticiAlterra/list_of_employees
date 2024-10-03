import os.path

import pandas as pd
def rec():
    path = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Рабочие", "Кадровая документация", "Рекомендации", "РЕКОМЕНДАЦИИ.xlsx")
    file_path_employees = os.sep * 2 + os.path.join("tg-storage01", "Служба персонала", "Общие", "Кадровый учет", "Действующие сотрудники.xlsx")

    # загрузка данных
    df = pd.read_excel(path, header=1)

    # замена пропущенных значений на пустые строки
    df['Рекомендации'] = df['Рекомендации'].fillna('')
    # заполнение пропущенных значений в столбце ФИО
    df['ФИО'] = df['ФИО'].ffill()
    # убираем лишние пробелы
    df['ФИО'] = df['ФИО'].str.strip()

    # группировка данных и объединение строк в столбце Рекомендации
    # объединение строк в столбце Рекомендации
    df_rec = df.groupby('ФИО')['Рекомендации'].apply(lambda x: '\n'.join(x)).reset_index()


    # # сохранение данных в файл Excel
    df_emp = pd.read_excel(file_path_employees)

    # Удаляем дубликаты по file_name
    df_rec.drop_duplicates(subset=['ФИО'], inplace=True)

    merged_df = pd.merge(df_emp, df_rec, left_on='ФИО', right_on='ФИО',
                         how='left')
    merged_df.to_excel(file_path_employees, index=False)
