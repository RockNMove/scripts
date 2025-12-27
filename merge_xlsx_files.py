
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog


def file_window():
    # Создаём скрытое окно
    root = tk.Tk()
    root.withdraw()  # Скрываем основное окно

    # Открываем диалог выбора файла
    file_path = filedialog.askdirectory(title="Выберите файл")

    return file_path


# Укажи путь к папке с Excel-файлами
folder_path = file_window()  # замени на свой путь
folder_name = os.path.basename(os.path.normpath(folder_path))

# Находим все Excel-файлы в папке
excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]

# Загружаем все таблицы, собираем уникальные колонки
dataframes = []
all_columns = set()

for file in excel_files:
    file_path = os.path.join(folder_path, file)
    try:
        df = pd.read_excel(file_path)
        df['source_file'] = file  # добавим колонку с именем файла
        dataframes.append(df)
        all_columns.update(df.columns)
    except Exception as e:
        print(f'⚠️ Ошибка при чтении файла {file}: {e}')

# Подготовим список отфильтрованных и приведённых таблиц
all_columns = list(all_columns)  # приводим к списку
reindexed_dfs = []

for df in dataframes:
    df = df.reindex(columns=all_columns, fill_value=pd.NA)
    if not df.empty and df.notna().any().any():  # исключаем полностью пустые
        reindexed_dfs.append(df)

# Объединяем все таблицы
if reindexed_dfs:
    merged_data = pd.concat(reindexed_dfs, ignore_index=True)
    output_file = os.path.join(folder_path, f'{folder_name}_merged_result.xlsx')
    merged_data.to_excel(output_file, index=False)
    print(f'✅ Объединённый файл сохранён: {output_file}')
else:
    print('❗Нет подходящих данных для объединения.')
