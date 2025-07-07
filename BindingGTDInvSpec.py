import os
import re
from PyPDF2 import PdfMerger


# --- Функция для извлечения номера из имени файла GTD ---
def get_gtd_number(file_name):
    """
    Извлекает последний числовой идентификатор из имени файла GTD.
    Пример: из "GTD_10228010_290625_5196376.pdf" вернет 5196376.
    """
    try:
        # Убираем расширение .pdf и разделяем по '_'
        parts = re.split(r'[_-]', file_name.replace('.pdf', ''))

        # Находим последний числовой элемент
        for part in reversed(parts):
            if part.isdigit():
                return int(part)
        return float('inf')  # Если число не найдено, отправить его в конец списка
    except Exception:
        return float('inf')


# --- Функция для извлечения номера из имени файла Invoice ---
def get_invoice_number(file_name):
    """
    Извлекает номер из имени файла Invoice.
    Пример: из "Invoice 1764.pdf" вернет 1764.
    """
    try:
        match = re.search(r'Invoice (\d+)', file_name)
        if match:
            return int(match.group(1))
        return float('inf')
    except Exception:
        return float('inf')


# --- Функция для чтения путей и диапазона из path.txt ---
def read_paths_and_range(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = [line.strip() for line in f if line.strip()]
            if len(lines) < 3:
                raise ValueError(
                    "Файл path.txt должен содержать три строки: путь к исходной папке, путь для сохранения и диапазон номеров папок."
                )
            source_path = lines[0]
            save_path = lines[1]
            folder_range_str = lines[2]
            return source_path, save_path, folder_range_str
    except FileNotFoundError:
        print(f"Ошибка: Файл '{file_path}' не найден. Убедитесь, что он существует.")
        return None, None, None
    except Exception as e:
        print(f"Ошибка при чтении path.txt: {e}")
        return None, None, None


# --- Функция для парсинга диапазона номеров папок ---
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


# --- Главная функция скрипта ---
def main():
    script_directory = os.path.dirname(os.path.realpath(__file__))
    path_file = os.path.join(script_directory, 'path.txt')

    source_path, save_path, folder_range_str = read_paths_and_range(path_file)

    if not (source_path and save_path and folder_range_str):
        print("Ошибка: Не удалось получить необходимые данные из path.txt. Завершение.")
        return

    # Проверяем существование исходной и целевой папок
    if not os.path.isdir(source_path):
        print(f"Ошибка: Исходная папка '{source_path}' не найдена или недоступна. Завершение.")
        return
    if not os.path.exists(save_path):
        os.makedirs(save_path)
        print(f"Создана папка для сохранения: {save_path}")

    valid_folders = parse_folder_range(folder_range_str)
    if not valid_folders:
        print("Не найдено номеров папок для обработки в диапазоне. Завершение.")
        return

    # Регулярное выражение для извлечения номера папки (в начале имени)
    folder_number_pattern = re.compile(r'^\d+')

    # Словарь для хранения пар GTD-Invoice: {GTD_номер_документа: {'gtd': 'путь_к_файлу_gtd.pdf', 'invoice': 'путь_к_файлу_invoice.pdf'}}
    paired_documents = {}

    # Проход по всем папкам в директории
    all_folder_names = os.listdir(source_path)
    # Сортируем папки по номеру, чтобы обрабатывать их последовательно
    sorted_folders = sorted(all_folder_names,
                            key=lambda name: int(
                                folder_number_pattern.match(name).group()) if folder_number_pattern.match(
                                name) else float('inf'))

    for folder_name in sorted_folders:
        folder_path = os.path.join(source_path, folder_name)

        if os.path.isdir(folder_path):
            folder_number_match = folder_number_pattern.search(folder_name)
            if not folder_number_match:
                print(f"Пропущена папка '{folder_name}' (не удалось извлечь номер).")
                continue

            current_folder_number = int(folder_number_match.group())

            # Проверяем, входит ли папка в указанный диапазон
            if current_folder_number in valid_folders:
                print(f"\nОбработка папки: {folder_name}")
                gtd_file_in_folder = None
                invoice_file_in_folder = None

                for file_name in os.listdir(folder_path):
                    file_lower = file_name.lower()
                    if file_lower.endswith(".pdf"):
                        if file_lower.startswith("gtd_"):
                            gtd_file_in_folder = os.path.join(folder_path, file_name)
                        elif "invoice" in file_lower:
                            invoice_file_in_folder = os.path.join(folder_path, file_name)

                # Если найдены GTD и Invoice в папке, связываем их
                if gtd_file_in_folder and invoice_file_in_folder:
                    gtd_num = get_gtd_number(os.path.basename(gtd_file_in_folder))
                    # Используем номер GTD как ключ для связывания
                    if gtd_num != float('inf'):
                        paired_documents[gtd_num] = {
                            'gtd': gtd_file_in_folder,
                            'invoice': invoice_file_in_folder
                        }
                        print(
                            f"Найдена пара GTD ({os.path.basename(gtd_file_in_folder)}) и Invoice ({os.path.basename(invoice_file_in_folder)})")
                    else:
                        print(
                            f"Предупреждение: Не удалось извлечь номер GTD из файла '{os.path.basename(gtd_file_in_folder)}'. Пара не будет добавлена.")
                elif gtd_file_in_folder:
                    print(f"В папке '{folder_name}' найден GTD, но нет Invoice. Пропускаем Invoice.")
                    gtd_num = get_gtd_number(os.path.basename(gtd_file_in_folder))
                    if gtd_num != float('inf'):
                        paired_documents[gtd_num] = {'gtd': gtd_file_in_folder, 'invoice': None}
                elif invoice_file_in_folder:
                    print(f"В папке '{folder_name}' найден Invoice, но нет GTD. Пропускаем GTD.")
                    # Если нужно добавить Invoice без GTD, можно раскомментировать и добавить логику.
                    # Сейчас Invoice без GTD не будет обработан.
                else:
                    print(f"В папке '{folder_name}' не найдено ни GTD, ни Invoice PDF-файлов. Пропускаем.")
            else:
                print(f"Пропущена папка '{folder_name}' (номер {current_folder_number} не входит в диапазон).")

    # Сортируем ключи (номера GTD) и формируем итоговый список файлов для объединения
    files_to_merge = []
    set_of_processed_folder_numbers = set() # Используем set для автоматического удаления дубликатов

    # Сортируем пары по номеру GTD
    sorted_gtd_numbers = sorted(paired_documents.keys())

    for gtd_num in sorted_gtd_numbers:
        pair = paired_documents[gtd_num]
        if pair['gtd']:
            files_to_merge.append(pair['gtd'])
            folder_num = get_folder_number_from_path(pair['gtd'])
            if folder_num != float('inf'):
                set_of_processed_folder_numbers.add(folder_num) # Добавляем номер папки в set

        if pair['invoice']:
            files_to_merge.append(pair['invoice'])
            # Мы не добавляем номер папки от Invoice, чтобы избежать дублирования или путаницы,
            # так как они уже связаны с GTD-номером.

    if not files_to_merge:
        print("Не найдено подходящих PDF-файлов (GTD или Invoice) для объединения в указанных папках.")
        return

    # Формирование диапазона номеров обработанных папок для имени файла
    # Преобразуем set в отсортированный список
    processed_folder_numbers_list = sorted(list(set_of_processed_folder_numbers))
    
    range_parts = []
    if processed_folder_numbers_list:
        current_range = [processed_folder_numbers_list[0]]
        for i in range(1, len(processed_folder_numbers_list)):
            if processed_folder_numbers_list[i] == processed_folder_numbers_list[i - 1] + 1:
                current_range.append(processed_folder_numbers_list[i])
            else:
                if len(current_range) > 1:
                    range_parts.append(f"{current_range[0]}-{current_range[-1]}")
                else:
                    range_parts.append(str(current_range[0]))
                current_range = [processed_folder_numbers_list[i]]

        if len(current_range) > 1:
            range_parts.append(f"{current_range[0]}-{current_range[-1]}")
        else:
            range_parts.append(str(current_range[0]))

    condensed_range_str = ';'.join(range_parts) if range_parts else "NoRange"

    # Создание имени выходного файла
    # --- ИСПРАВЛЕНИЕ ЗДЕСЬ: Добавляем точку в конце "pcs." ---
    output_file_name = f"GTD+Invoice {condensed_range_str} {len(processed_folder_numbers_list)} pcs..pdf"
    output_file_path = os.path.join(save_path, output_file_name)

    # Объединение PDF-файлов
    merger = PdfMerger()
    print("\nНачинаю объединение файлов...")

    for pdf_path in files_to_merge:
        print(f"-> Добавляю: {os.path.basename(pdf_path)}")
        try:
            with open(pdf_path, 'rb') as f:
                merger.append(f)
        except Exception as e:
            print("\n" + "=" * 50)
            print(f"!!! КРИТИЧЕСКАЯ ОШИБКА ПРИ ДОБАВЛЕНИИ ФАЙЛА: '{pdf_path}'")
            print(f"!!! Текст ошибки: {e}")
            print("!!! Этот файл будет пропущен. Возможно, он поврежден или заблокирован.")
            print("=" * 50 + "\n")
            continue

    print("\nВсе файлы успешно добавлены в объект. Пытаюсь сохранить на диск...")
    try:
        with open(output_file_path, 'wb') as f_out:
            merger.write(f_out)

        merger.close()

        # --- ИСПРАВЛЕНИЕ ЗДЕСЬ: Меняем вывод в консоль для отображения количества папок ---
        print(f"\nОбъединено {len(files_to_merge)} файлов из {len(processed_folder_numbers_list)} папок.")
        print(f"Объединённый файл сохранён как: {output_file_name}")

    except Exception as e:
        print("\n" + "=" * 50)
        print(f"!!! ОШИБКА НА ЭТАПЕ ЗАПИСИ ИТОГОВОГО ФАЙЛА!")
        print(
            f"!!! Это означает, что проблема с правами доступа к папке '{save_path}' или один из добавленных файлов вызвал внутреннюю ошибку PyPDF2.")
        print(f"!!! Текст ошибки: {e}")
        print("=" * 50 + "\n")


def get_folder_number_from_path(file_path):
    """
    Извлекает номер папки из полного пути к файлу.
    Предполагает, что номер папки находится в начале имени папки.
    """
    folder_name = os.path.basename(os.path.dirname(file_path))
    match = re.compile(r'^\d+').match(folder_name)
    return int(match.group()) if match else float('inf')


if __name__ == "__main__":
    main()
