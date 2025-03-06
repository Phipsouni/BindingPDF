import os
import re
from PyPDF2 import PdfMerger
import sys

# Чтение путей и диапазона из path.txt
with open('path.txt', 'r') as file:
    paths = file.readlines()
    source_path = paths[0].strip()  # Путь к директории с папками
    save_path = paths[1].strip()  # Путь для сохранения скрепленного файла
    folder_range = paths[2].strip()  # Диапазон номеров папок

# Парсинг диапазона папок
def parse_folder_range(range_str):
    ranges = range_str.split(',')
    folder_numbers = set()
    for r in ranges:
        if '-' in r:
            start, end = map(int, r.split('-'))
            folder_numbers.update(range(start, end + 1))
        else:
            folder_numbers.add(int(r))
    return sorted(folder_numbers)

valid_folders = parse_folder_range(folder_range)

# Регулярное выражение для извлечения первого числа в названии папки
folder_number_pattern = re.compile(r'^\d+')

# Список номеров папок, которые были обработаны
processed_folders = []

# Массив для хранения файлов из всех папок
all_pdfs = []

# Сортируем список папок по первому числу в названии
def get_folder_number(folder_name):
    match = folder_number_pattern.match(folder_name)
    return int(match.group()) if match else float('inf')  # Если число не найдено, отправить папку в конец

sorted_folders = sorted(os.listdir(source_path), key=get_folder_number)

# Проход по всем папкам в заданной директории
for folder_name in sorted_folders:
    folder_path = os.path.join(source_path, folder_name)

    # Проверка на то, что это папка
    if os.path.isdir(folder_path):
        # Извлечение номера папки
        folder_number_match = folder_number_pattern.search(folder_name)
        if folder_number_match:
            folder_number = int(folder_number_match.group())
        else:
            continue

        # Проверяем, входит ли номер папки в заданный диапазон
        if folder_number in valid_folders:
            gtd_files = []
            esd_files = []

            # Проход по файлам в папке
            for file_name in os.listdir(folder_path):
                if file_name.lower().endswith(".pdf"):  # Только PDF
                    file_path = os.path.join(folder_path, file_name)
                    if file_name.startswith("GTD_"):
                        gtd_files.append(file_path)
                    else:
                        esd_files.append(file_path)

            # Пропуск папки, если нет GTD-файлов
            if not gtd_files:
                print(f"Пропущена папка {folder_name} (нет GTD-файлов).")
                continue

            processed_folders.append(folder_number)

            # Добавляем файлы в порядке: сначала GTD, затем ЭСД
            all_pdfs.extend(sorted(gtd_files)[:1])  # GTD файл (первый найденный)
            all_pdfs.extend(sorted(esd_files)[:1])  # ЭСД файл (первый найденный)

# Упорядочивание номеров обработанных папок
processed_folders.sort()

# Генерация диапазона номеров папок для имени файла
range_parts = []
current_range = [processed_folders[0]]

for i in range(1, len(processed_folders)):
    if processed_folders[i] == processed_folders[i - 1] + 1:
        current_range.append(processed_folders[i])
    else:
        if len(current_range) > 1:
            range_parts.append(f"{current_range[0]}-{current_range[-1]}")
        else:
            range_parts.append(str(current_range[0]))
        current_range = [processed_folders[i]]

if len(current_range) > 1:
    range_parts.append(f"{current_range[0]}-{current_range[-1]}")
else:
    range_parts.append(str(current_range[0]))

range_str = ';'.join(range_parts)

# Создание итогового имени файла
output_file_name = f"GTD+ЭСД {range_str} {len(processed_folders)} pcs..pdf"
output_file_path = os.path.join(save_path, output_file_name)

# Объединение PDF-файлов
if all_pdfs:
    merger = PdfMerger()
    for pdf in all_pdfs:
        merger.append(pdf)

    # Сохранение объединённого PDF
    merger.write(output_file_path)
    merger.close()

    print(f"Объединённый файл сохранён как: {output_file_name}")
else:
    print("Не найдено файлов для объединения в указанном диапазоне папок.")
