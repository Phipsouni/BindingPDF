import os
import re
from PyPDF2 import PdfMerger


# --- НОВАЯ ФУНКЦИЯ для извлечения номера Invoice из имени файла ---
def get_invoice_number(file_name):
    """
    Извлекает номер из имени файла Invoice.
    Пример: из "Invoice 1764.pdf" вернет 1764.
    """
    try:
        # Используем регулярное выражение для поиска чисел после "Invoice"
        match = re.search(r'Invoice (\d+)', file_name, re.IGNORECASE)  # re.IGNORECASE для "invoice", "Invoice" и т.д.
        if match:
            return int(match.group(1))
        return float('inf')  # Если число не найдено, отправить его в конец списка
    except Exception:
        # Если формат имени файла неверный, отправляем его в конец списка
        return float('inf')


# --- Чтение путей и диапазона из path.txt ---
# Эта часть остаётся без изменений, так как формат path.txt тот же
def read_paths_and_range(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = [line.strip() for line in f if line.strip()]
            if len(lines) < 3:
                raise ValueError(
                    "Файл path.txt должен содержать три строки: путь к исходной папке, путь для сохранения и диапазон номеров папок.")
            source_path = lines[0]
            save_path = lines[1]
            folder_range = lines[2]
            return source_path, save_path, folder_range
    except FileNotFoundError:
        print(f"Ошибка: Файл '{file_path}' не найден. Убедитесь, что он существует.")
        return None, None, None
    except Exception as e:
        print(f"Ошибка при чтении path.txt: {e}")
        return None, None, None


# --- Функция для парсинга диапазона номеров папок ---
# Эта часть остаётся без изменений
def parse_folder_range(range_str):
    ranges = range_str.split(',')
    folder_numbers = set()
    for r in ranges:
        r = r.strip()
        if '-' in r:
            try:
                start, end = map(int, r.split('-'))
                if start > end:
                    print(f"Предупреждение: Неверный диапазон '{r}'. Начало должно быть меньше или равно концу.")
                    continue
                folder_numbers.update(range(start, end + 1))
            except ValueError:
                print(f"Предупреждение: Неверный формат диапазона '{r}'. Пропускаем.")
        else:
            try:
                folder_numbers.add(int(r))
            except ValueError:
                print(f"Предупреждение: Неверный формат числа '{r}'. Пропускаем.")
    return sorted(list(folder_numbers))


# --- Функция для извлечения номера папки ---
# Эта часть остаётся без изменений
def get_folder_number(folder_name):
    match = re.compile(r'^\d+').match(folder_name)
    return int(match.group()) if match else float('inf')


# --- Главный блок выполнения скрипта ---
if __name__ == "__main__":
    script_directory = os.path.dirname(os.path.realpath(__file__))
    path_file = os.path.join(script_directory, 'path.txt')

    source_path, save_path, folder_range_str = read_paths_and_range(path_file)

    if not (source_path and save_path and folder_range_str):
        print("Ошибка: Не удалось получить необходимые данные из path.txt. Завершение.")
        exit()  # Используем exit() для немедленного завершения

    # Проверяем существование исходной и целевой папок
    if not os.path.isdir(source_path):
        print(f"Ошибка: Исходная папка '{source_path}' не найдена или недоступна. Завершение.")
        exit()
    if not os.path.exists(save_path):
        os.makedirs(save_path)
        print(f"Создана папка для сохранения: {save_path}")

    valid_folders = parse_folder_range(folder_range_str)
    if not valid_folders:
        print("Не найдено номеров папок для обработки в диапазоне. Завершение.")
        exit()

    # Регулярное выражение для извлечения номера папки
    folder_number_pattern = re.compile(r'^\d+')

    # Список для хранения ВСЕХ найденных Invoice PDF-файлов
    all_invoice_pdfs = []
    processed_folders_numbers = []  # Список номеров папок, которые были обработаны (для формирования имени файла)

    # Получаем и сортируем папки по номеру
    all_folder_names = os.listdir(source_path)
    sorted_folders = sorted(all_folder_names, key=get_folder_number)

    print(f"Начинаем поиск файлов 'Invoice.pdf' в папках из диапазона: {valid_folders}")

    # Проход по всем папкам в директории
    for folder_name in sorted_folders:
        folder_path = os.path.join(source_path, folder_name)

        if os.path.isdir(folder_path):
            folder_number_match = folder_number_pattern.search(folder_name)
            if not folder_number_match:
                # print(f"Пропущена папка '{folder_name}' (не удалось извлечь номер).") # Закомментировано для уменьшения шума
                continue

            current_folder_number = int(folder_number_match.group())

            # Проверяем, входит ли папка в указанный диапазон
            if current_folder_number in valid_folders:
                invoice_files_in_folder = []

                # Проход по файлам в папке
                for file_name in os.listdir(folder_path):
                    # Ищем файлы, содержащие "Invoice" (без учета регистра) и заканчивающиеся на ".pdf"
                    if "invoice" in file_name.lower() and file_name.lower().endswith(".pdf"):
                        file_path = os.path.join(folder_path, file_name)
                        invoice_files_in_folder.append(file_path)

                # Если в папке нашлись Invoice-файлы
                if invoice_files_in_folder:
                    # Если найдено несколько Invoice в одной папке, это может быть проблемой.
                    # В этом скрипте мы добавим все найденные Invoice. Если нужен только один,
                    # можно добавить условие типа: `if len(invoice_files_in_folder) > 1: print("Warning...")`
                    all_invoice_pdfs.extend(invoice_files_in_folder)
                    processed_folders_numbers.append(current_folder_number)
                    print(f"Найдено Invoice PDF в папке '{folder_name}'.")
                else:
                    print(f"Пропущена папка '{folder_name}' (нет файлов Invoice PDF).")
            # else: # Закомментировано для уменьшения шума, если папка вне диапазона
            #     print(f"Пропущена папка '{folder_name}' (номер {current_folder_number} не входит в диапазон).")

    # --- ИЗМЕНЕНИЕ: Сортируем ВЕСЬ список Invoice файлов по номеру Invoice ---
    # Эта сортировка выполняется после того, как все файлы из всех папок собраны
    all_invoice_pdfs.sort(key=lambda x: get_invoice_number(os.path.basename(x)))

    # Формирование диапазона номеров обработанных папок для имени выходного файла
    processed_folders_numbers = sorted(list(set(processed_folders_numbers)))  # Удаляем дубликаты и сортируем
    range_parts = []
    if processed_folders_numbers:
        current_range = [processed_folders_numbers[0]]
        for i in range(1, len(processed_folders_numbers)):
            if processed_folders_numbers[i] == processed_folders_numbers[i - 1] + 1:
                current_range.append(processed_folders_numbers[i])
            else:
                if len(current_range) > 1:
                    range_parts.append(f"{current_range[0]}-{current_range[-1]}")
                else:
                    range_parts.append(str(current_range[0]))
                current_range = [processed_folders_numbers[i]]

        if len(current_range) > 1:
            range_parts.append(f"{current_range[0]}-{current_range[-1]}")
        else:
            range_parts.append(str(current_range[0]))

    condensed_range_str = ';'.join(range_parts) if range_parts else "NoRange"

    # Создание имени выходного файла
    output_file_name = f"Inv. + Spec. {condensed_range_str} {len(all_invoice_pdfs)} pcs..pdf"
    output_file_path = os.path.join(save_path, output_file_name)

    # Объединение PDF-файлов
    if all_invoice_pdfs:
        merger = PdfMerger()
        print("\nНачинаю объединение файлов...")

        for pdf_path in all_invoice_pdfs:
            print(f"-> Добавляю: {os.path.basename(pdf_path)}")
            try:
                with open(pdf_path, 'rb') as f:
                    merger.append(f)
            except Exception as e:
                print("\n" + "=" * 50)
                print(f"!!! КРИТИЧЕСКАЯ ОШИБКА ПРИ ДОБАВЛЕНИИ ФАЙЛА:")
                print(f"!!! Файл: {pdf_path}")
                print(f"!!! Текст ошибки: {e}")
                print("!!! Этот файл будет пропущен. Возможно, он поврежден или заблокирован.")
                print("=" * 50 + "\n")
                continue  # Продолжаем даже если один файл поврежден

        print("\nВсе файлы успешно добавлены в объект. Пытаюсь сохранить на диск...")
        try:
            with open(output_file_path, 'wb') as f_out:
                merger.write(f_out)

            merger.close()

            print(f"\nНайдено и объединено {len(all_invoice_pdfs)} файлов.")
            print(f"Объединённый файл сохранён как: '{output_file_name}' в папке '{save_path}'")

        except Exception as e:
            print("\n" + "=" * 50)
            print(f"!!! ОШИБКА НА ЭТАПЕ ЗАПИСИ ФАЙЛА!")
            print(
                f"!!! Это означает, что проблема с правами доступа к папке '{save_path}' или один из добавленных файлов вызвал внутреннюю ошибку PyPDF2.")
            print(f"!!! Текст ошибки: {e}")
            print("=" * 50 + "\n")

    else:
        print("Не найдено Invoice PDF-файлов для объединения в указанном диапазоне папок.")
