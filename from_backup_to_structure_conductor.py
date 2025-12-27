from pathlib import Path
import csv


def extract_bases(file_path, set_to_extract, first=True):

    if not set_to_extract:
        return set_to_extract

    p = Path(file_path)
    with p.open(encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=';')

        level_set = set()
        for row in reader:
            for item in set_to_extract:
                if row['База'] == item:
                    if first:
                        level_set.add(item)

                    base_name = row['Название'].split(
                        ' / ')[0]

                    level_set.add(base_name)

        return level_set.union(extract_bases(file_path, level_set, first=False))


def make_main_list(file_path, bases_set):

    p = Path(file_path)
    with p.open(encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=';')

        lst = []
        for row in reader:
            for item in bases_set:
                if row['База'] == item:
                    lst.append(row)

        empty_cols = {col for col in reader.fieldnames if all(not dict[col] for dict in lst)}

        for dic in lst:
            for col in empty_cols:
                dic.pop(col)

        return lst


def make_spec_list(file_path, bases_set):

    p = Path(file_path)
    with p.open(encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=';')

        lst = []
        for row in reader:
            for item in bases_set:
                if row['База'] == item:
                    lst.append(row)

        empty_cols = {col for col in reader.fieldnames if all(not dict[col] for dict in lst)}

        for dic in lst:
            for col in empty_cols:
                dic.pop(col)

        return lst


def make_tech_list(file_path, bases_set):

    p = Path(file_path)
    with p.open(encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=';')

        lst = []
        for row in reader:
            for item in bases_set:
                if row['База'] == item:
                    lst.append(row)

        empty_cols = {col for col in reader.fieldnames if all(not dict[col] for dict in lst)}

        for dic in lst:
            for col in empty_cols:
                dic.pop(col)

        return lst


def make_units_list(file_path, full_units_set):

    p = Path(file_path)
    with p.open(encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=';')

        lst = []
        for row in reader:
            for item in full_units_set:
                if row['Название'] == item:
                    lst.append(row)

        empty_cols = {col for col in reader.fieldnames if all(not dict[col] for dict in lst)}

        for dic in lst:
            if any(item in dic['Категории'] for item in ['ПКИ', 'ТИ']):
                dic['Производство'] = 'нет'
                dic['Закупка'] = 'да'
            else:
                dic['Производство'] = 'да'
                dic['Закупка'] = 'нет'

            for col in empty_cols:
                dic.pop(col)
                dic['Категория'] = dic.pop('Категории')

        return lst


# Менять тут:
#################################################################################################################
set_to_extract = {'Шнек', 'Удлинитель шнека', 'Шнек универсальный'}

main_file_path = r"C:\Users\Andrey\Desktop\проливка_хтк\CSV_Основное_merged_result_10.10.25.csv"
spec_file_path = r"C:\Users\Andrey\Desktop\проливка_хтк\CSV_Спецификации_merged_result_10.10.25.csv"
tech_file_path = r"C:\Users\Andrey\Desktop\проливка_хтк\CSV_Техпроцесс_merged_result_10.10.25.csv"
units_file_path = r"C:\Users\Andrey\Desktop\проливка_хтк\CSV_полный_справочник_10.10.25.csv"

aim_main_file_path = r"C:\Users\Andrey\Desktop\проливка_хтк\New_CSV_Основное.csv"
aim_spec_file_path = r"C:\Users\Andrey\Desktop\проливка_хтк\New_CSV_Спецификации.csv"
aim_tech_file_path = r"C:\Users\Andrey\Desktop\проливка_хтк\New_CSV_Техпроцесс.csv"
aim_units_file_path = r"C:\Users\Andrey\Desktop\проливка_хтк\New_CSV_Юниты.csv"
#################################################################################################################

bases_set = extract_bases(spec_file_path, set_to_extract)

main_dicts_list = make_main_list(main_file_path, bases_set)  # лист словарей к записи в файл
spec_dicts_list = make_spec_list(spec_file_path, bases_set)  # лист словарей к записи в файл
tech_dicts_list = make_tech_list(tech_file_path, bases_set)  # лист словарей к записи в файл

full_units_set = {dict['Название'] for dict in main_dicts_list}.union(
    {dict['Название'] for dict in spec_dicts_list})

units_dicts_list = make_units_list(units_file_path, full_units_set)  # лист словарей к записи в файл

print(f"{len(bases_set)} pcs bases set is:\n {bases_set}")
print()
print(f"{len(full_units_set)} pcs full units set is:\n {full_units_set}")
print()

p = Path(aim_main_file_path)
with p.open('w', newline='', encoding='utf-8-sig') as f:

    headers = main_dicts_list[0].keys()
    writer = csv.DictWriter(f, delimiter=';', fieldnames=headers)
    writer.writeheader()
    writer.writerows(main_dicts_list)
    print(f"{p.name} for bases {set_to_extract} extracted to {p.parent}")

p = Path(aim_spec_file_path)
with p.open('w', newline='', encoding='utf-8-sig') as f:

    headers = spec_dicts_list[0].keys()
    writer = csv.DictWriter(f, delimiter=';', fieldnames=headers)
    writer.writeheader()
    writer.writerows(spec_dicts_list)
    print(f"{p.name} for bases {set_to_extract} extracted to {p.parent}")

p = Path(aim_tech_file_path)
with p.open('w', newline='', encoding='utf-8-sig') as f:

    headers = tech_dicts_list[0].keys()
    writer = csv.DictWriter(f, delimiter=';', fieldnames=headers)
    writer.writeheader()
    writer.writerows(tech_dicts_list)
    print(f"{p.name} for bases {set_to_extract} extracted to {p.parent}")

p = Path(aim_units_file_path)
with p.open('w', newline='', encoding='utf-8-sig') as f:

    headers = units_dicts_list[0].keys()
    writer = csv.DictWriter(f, delimiter=';', fieldnames=headers)
    writer.writeheader()
    writer.writerows(units_dicts_list)
    print(f"{p.name} for bases {set_to_extract} extracted to {p.parent}")
