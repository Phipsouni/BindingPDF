import os
import re
from PyPDF2 import PdfMerger


# --- НОВАЯ ФУНКЦИЯ для извлечения номера документа из имени файла ---
def get_document_number(file_path):
    """
    Извлекает последний числовой идентификатор из имени файла.
    Пример: из "GTD_10228010_290625_5196376.pdf" вернет 5196376.
    """
    try:
        # Получаем только имя файла из полного пути
        file_name = os.path.basename(file_path)
        # Убираем расширение .pdf и разделяем по '_'
        # Используем re.split для обработки случаев, когда номер может быть в конце без подчеркивания,
        # хотя в примере "GTD_10228010_290625_5196376.pdf" это не так
        parts = re.split(r'[_-]', file_name.replace('.pdf', ''))
        
        # Находим последний числовой элемент
        for part in reversed(parts):
            if part.isdigit():
                return int(part)
        return float('inf') # Если число не найдено, отправить его в конец списка
    except Exception:
        # Если формат имени файла неверный, отправляем его в конец списка
        return float('inf')


# Чтение путей и диапазона из path.txt
with open('path.txt', 'r', encoding='utf-8') as file:
    paths = file.readlines()
    source_path = paths[0].strip()  # Путь к директории с папками
    save_path = paths[1].strip()  # Путь для сохранения скрепленного файла
    folder_range = paths[2].strip()  # Диапазон номеров папок


# Функция для парсинга диапазона номеров папок
def parse_folder_range(range_str):
    ranges = range_str.split(',')
    folder_numbers = set()
    for r in ranges:
        if '-' in r:
            start, end = map(int, r.split('-'))
            folder_numbers.update(range(start, end + 1))
        else:
            folder_numbers.add(int(r))
    return sorted(list(folder_numbers)) # Преобразуем в список для сортировки


valid_folders = parse_folder_range(folder_range)

# Регулярное выражение для извлечения номера папки
folder_number_pattern = re.compile(r'^\d+')


# Функция для сортировки папок по номеру
def get_folder_number(folder_name):
    match = folder_number_pattern.match(folder_name)
    return int(match.group()) if match else float('inf')  # Если число не найдено, отправить в конец


sorted_folders = sorted(os.listdir(source_path), key=get_folder_number)

# Список обработанных папок
processed_folders = []

# Список для хранения ВСЕХ найденных GTD-файлов
all_pdfs = []

# Проход по всем папкам в директории
for folder_name in sorted_folders:
    folder_path = os.path.join(source_path, folder_name)

    # Проверяем, является ли объект папкой
    if os.path.isdir(folder_path):
        # Извлекаем номер папки
        folder_number_match = folder_number_pattern.search(folder_name)
        if folder_number_match:
            folder_number = int(folder_number_match.group())
        else:
            continue

        # Проверяем, входит ли папка в указанный диапазон
        if folder_number in valid_folders:
            gtd_files_in_folder = []

            # Проход по файлам в папке
            for file_name in os.listdir(folder_path):
                if file_name.lower().endswith(".pdf") and file_name.startswith("GTD_"):
                    file_path = os.path.join(folder_path, file_name)
                    gtd_files_in_folder.append(file_path)

            # Если в папке нашлись GTD-файлы, добавляем их в общий список
            if gtd_files_in_folder:
                processed_folders.append(folder_number)
                all_pdfs.extend(gtd_files_in_folder)
            else:
                print(f"Пропущена папка {folder_name} (нет файлов GTD_).")

# --- ИЗМЕНЕНИЕ: Сортируем ВЕСЬ список файлов по номеру документа ---
# Эта сортировка выполняется после того, как все файлы из всех папок собраны
all_pdfs.sort(key=get_document_number)

# Формирование диапазона номеров обработанных папок
processed_folders.sort()
range_parts = []
if processed_folders:
    current_range = [processed_folders[0]]
    for i in range(1, len(processed_folders)):
        if processed_folders[i] == processed_folders[i - 1] + 1:
            if len(current_range) > 1:
                current_range.pop() # Удаляем предыдущий элемент, так как он часть диапазона
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

# Используем уже созданный список с диапазонами 'range_parts'
condensed_range_str = ';'.join(range_parts)

# Создание имени выходного файла
output_file_name = f"GTD {condensed_range_str} {len(all_pdfs)} pcs.pdf"
output_file_path = os.path.join(save_path, output_file_name) # *** ИСПРАВЛЕНИЕ: Определяем полный путь ***

# Объединение PDF-файлов
if all_pdfs:
    merger = PdfMerger()
    print("Начинаю объединение файлов...")

    # Этот цикл попытается добавить все файлы и сообщит, если какой-то вызовет ошибку сразу
    for pdf_path in all_pdfs:
        print(f"-> Добавляю в очередь: {os.path.basename(pdf_path)}")
        try:
            # Открываем файл в бинарном режиме для большей надежности
            with open(pdf_path, 'rb') as f:
                merger.append(f)
        except Exception as e:
            print("\n" + "="*50)
            print(f"!!! КРИТИЧЕСКАЯ ОШИБКА ПРИ ДОБАВЛЕНИИ ФАЙЛА:")
            print(f"!!! Файл: {pdf_path}")
            print(f"!!! Текст ошибки: {e}")
            print("="*50 + "\n")
            # Если нужно, можно раскомментировать, чтобы пропустить битый файл
            # continue  

    # Если все файлы добавились без ошибок, пытаемся сохранить
    print("\nВсе файлы успешно добавлены в объект. Пытаюсь сохранить на диск...")
    try:
        # Сохранение объединенного PDF
        with open(output_file_path, 'wb') as f_out:
            merger.write(f_out)
            
        merger.close()

        print(f"\nНайдено и объединено {len(all_pdfs)} файлов.")
        print(f"Объединённый файл сохранён как: {output_file_name}")

    except Exception as e:
        print("\n" + "="*50)
        print(f"!!! ОШИБКА НА ЭТАПЕ ЗАПИСИ ФАЙЛА!")
        print(f"!!! Это означает, что проблема в одном из ранее добавленных файлов или с правами доступа к папке.")
        print(f"!!! Текст ошибки: {e}")
        print("="*50 + "\n")

else:
    print("Не найдено GTD-файлов для объединения в указанном диапазоне папок.")
