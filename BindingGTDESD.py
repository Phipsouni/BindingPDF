import os
import re
from PyPDF2 import PdfMerger
import sys

# Чтение путей и диапазона из path.txt
try:
    with open('path.txt', 'r') as file:
        paths = file.readlines()
        if len(paths) < 3:
            print(
                "Ошибка: файл path.txt должен содержать 3 строки (путь к исходным файлам, путь для сохранения, диапазон папок).")
            sys.exit()
        source_path = paths[0].strip()
        save_path = paths[1].strip()
        folder_range = paths[2].strip()
except FileNotFoundError:
    print("Ошибка: файл 'path.txt' не найден. Пожалуйста, создайте его.")
    sys.exit()


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
    return int(match.group()) if match else float('inf')


# Проверка существования исходной директории
if not os.path.isdir(source_path):
    print(f"Ошибка: Исходная директория не найдена по пути: {source_path}")
    sys.exit()

sorted_folders = sorted(os.listdir(source_path), key=get_folder_number)

# Проход по всем папкам в заданной директории
for folder_name in sorted_folders:
    folder_path = os.path.join(source_path, folder_name)

    # Проверка на то, что это папка
    if os.path.isdir(folder_path):
        folder_number_match = folder_number_pattern.search(folder_name)
        if not folder_number_match:
            continue

        folder_number = int(folder_number_match.group())

        # Проверяем, входит ли номер папки в заданный диапазон
        if folder_number in valid_folders:
            gtd_files = []
            esd_files = []

            # Проход по файлам в папке
            for file_name in os.listdir(folder_path):
                if file_name.lower().endswith(".pdf"):
                    file_path = os.path.join(folder_path, file_name)

                    # --- ИЗМЕНЕНИЕ ЗДЕСЬ ---
                    # Если файл начинается с "GTD_", это GTD-файл
                    if file_name.startswith("GTD_"):
                        gtd_files.append(file_path)
                    # Если в имени файла ровно 4 дефиса, это ЭСД-файл
                    elif file_name.count('-') == 4:
                        esd_files.append(file_path)

            # Пропуск папки, если нет GTD или ЭСД файлов
            if not gtd_files:
                print(f"Пропущена папка {folder_name} (нет GTD-файлов).")
                continue
            if not esd_files:
                print(f"Пропущена папка {folder_name} (нет ЭСД-файлов с 4-мя дефисами).")
                continue

            processed_folders.append(folder_number)

            # Добавляем файлы в порядке: сначала GTD, затем ЭСД
            all_pdfs.extend(sorted(gtd_files)[:1])  # Первый найденный GTD файл
            all_pdfs.extend(sorted(esd_files)[:1])  # Первый найденный ЭСД файл

# Проверка, были ли найдены папки для обработки
if not processed_folders:
    print("Не найдено папок для обработки в указанном диапазоне.")
    sys.exit()

# Упорядочивание номеров обработанных папок
processed_folders.sort()

# Генерация диапазона номеров папок для имени файла
range_parts = []
if processed_folders:
    start_range = processed_folders[0]
    for i in range(1, len(processed_folders)):
        if processed_folders[i] != processed_folders[i - 1] + 1:
            end_range = processed_folders[i - 1]
            if start_range == end_range:
                range_parts.append(str(start_range))
            else:
                range_parts.append(f"{start_range}-{end_range}")
            start_range = processed_folders[i]

    # Добавляем последний диапазон
    end_range = processed_folders[-1]
    if start_range == end_range:
        range_parts.append(str(start_range))
    else:
        range_parts.append(f"{start_range}-{end_range}")

range_str = ';'.join(range_parts)

# Создание итогового имени файла
output_file_name = f"GTD+ЭСД {range_str} {len(processed_folders)} pcs..pdf"
output_file_path = os.path.join(save_path, output_file_name)

# Объединение PDF-файлов
if all_pdfs:
    merger = PdfMerger()
    for pdf in all_pdfs:
        try:
            merger.append(pdf)
        except Exception as e:
            print(f"Не удалось добавить файл {pdf}: {e}")
            continue

    # Сохранение объединённого PDF
    merger.write(output_file_path)
    merger.close()

    print(f"Объединённый файл сохранён как: {output_file_name}")
else:
    print("Не найдено PDF файлов для объединения в указанном диапазоне папок.")
