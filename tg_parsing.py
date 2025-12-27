import tkinter as tk
from tkinter import filedialog
import json
import re
import os
import datetime
import pandas as pd


# 🔁 Рекурсивная функция для извлечения текста из вложенных структур
def flatten_text(text_block):
    result = ""

    if isinstance(text_block, str):
        result += text_block
    elif isinstance(text_block, dict):
        result += text_block.get("text", "")
    elif isinstance(text_block, list):
        for item in text_block:
            result += flatten_text(item)

    return result


# 🧠 Основная функция: ищет по шаблонам
def extract_by_patterns(data, patterns):
    results = []

    for message in data.get("messages", []):
        text = message.get("text", [])
        full_text = flatten_text(text)

        entry = {
            "id": message.get("id"),
            "time": datetime.datetime.fromtimestamp(int(message.get("date_unixtime"))),
        }

        found = False

        for label, regex in patterns.items():
            match = re.search(regex, full_text)
            if match:
                try:
                    value = float(match.group(1))
                except ValueError:
                    value = match.group(1)  # если не число, просто строка
                entry[label] = value
                found = True

        if found:
            results.append(entry)

    return results


# 🧩 Добавляй нужные шаблоны ниже
search_patterns = {
    "bidask_depth_60": r"60%\s*—\s*([\d.]+)",
    "bidask_depth_8": r"8%\s*—\s*([\d.]+)",
    "bidask_depth_3": r"3%\s*—\s*([\d.]+)",
    "funding_high": r"Выше\s*—\s*([\d.]+)%",
    "funding_standard": r"Стандартно\s*—\s*([\d.]+)%",
    "funding_low": r"Ниже\s*—\s*([\d.]+)%",
    "demand_index": r"Индикатор спроса\s*=\s*([\d.]+)",

}

# 📦 выбор файла, чтение, запуск, сохранение
# выбор файла
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title="Выберите файл")
file_dir = os.path.dirname(file_path)

# чтение
with open(file_path, encoding="utf-8") as json_file:
    data = json.load(json_file)

# парсинг
output = extract_by_patterns(data, search_patterns)

# конвертация
json_output = json.dumps(output, ensure_ascii=False, indent=2, default=str)
df = pd.DataFrame(output)

# сохранение в:
json_txt_path = os.path.join(file_dir, "extracted_results.txt")
excel_path = os.path.join(file_dir, "extracted_results.xlsx")

# JSON
with open(json_txt_path, "w", encoding="utf-8") as f_out:
    f_out.write(json_output)

# XLSX
df.to_excel(excel_path, index=False)

print(f"Готово. Результат сохранён в: {file_dir}")
