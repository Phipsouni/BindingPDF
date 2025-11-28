import os
import re
import sys
from PyPDF2 import PdfMerger


# ==========================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (Утилиты)
# ==========================================

def get_clean_path(prompt_text):
    """Запрашивает путь у пользователя и удаляет кавычки, если они есть."""
    path = input(f"{prompt_text}: ").strip()
    # Удаляем кавычки в начале и конце, если они есть
    if (path.startswith('"') and path.endswith('"')) or (path.startswith("'") and path.endswith("'")):
        path = path[1:-1]
    return path


def parse_folder_range(range_str):
    """Парсит строку диапазона (например, '3550-3553,3560') в список чисел."""
    ranges = range_str.split(',')
    folder_numbers = set()
    for r in ranges:
        r = r.strip()
        if not r: continue
        if '-' in r:
            try:
                start, end = map(int, r.split('-'))
                if start > end:
                    print(f"Предупреждение: Неверный диапазон '{r}'.")
                    continue
                folder_numbers.update(range(start, end + 1))
            except ValueError:
                print(f"Предупреждение: Неверный формат диапазона '{r}'.")
        else:
            try:
                folder_numbers.add(int(r))
            except ValueError:
                print(f"Предупреждение: Неверный формат числа '{r}'.")
    return sorted(list(folder_numbers))


def get_folder_number(folder_name):
    """Извлекает номер папки для сортировки."""
    match = re.match(r'^\d+', folder_name)
    return int(match.group()) if match else float('inf')


def generate_range_string(processed_numbers):
    """Генерирует красивую строку диапазона (3550-3552;3560) для имени файла."""
    if not processed_numbers:
        return "NoRange"

    processed_numbers = sorted(list(set(processed_numbers)))
    range_parts = []
    current_range = [processed_numbers[0]]

    for i in range(1, len(processed_numbers)):
        if processed_numbers[i] == processed_numbers[i - 1] + 1:
            current_range.append(processed_numbers[i])
        else:
            if len(current_range) > 1:
                range_parts.append(f"{current_range[0]}-{current_range[-1]}")
            else:
                range_parts.append(str(current_range[0]))
            current_range = [processed_numbers[i]]

    if len(current_range) > 1:
        range_parts.append(f"{current_range[0]}-{current_range[-1]}")
    else:
        range_parts.append(str(current_range[0]))

    return ';'.join(range_parts)


def save_merged_pdf(merger, save_path, file_name):
    """Сохраняет PDF и обрабатывает ошибки."""
    full_path = os.path.join(save_path, file_name)
    try:
        if not os.path.exists(save_path):
            os.makedirs(save_path)

        print("\nСохранение файла на диск...")
        with open(full_path, 'wb') as f_out:
            merger.write(f_out)
        merger.close()
        print(f"✅ Успешно! Файл сохранён: {file_name}")
        print(f"📂 Путь: {save_path}")
    except Exception as e:
        print(f"❌ Ошибка при сохранении: {e}")


# ==========================================
# ЛОГИКА 1: BindingInvSpec (Инвойсы и Спецификации)
# ==========================================
def process_inv_spec(source_path, save_path, valid_folders):
    print("\n--- Запуск: Инвойсы и Спецификации ---")

    def get_invoice_num(fname):
        match = re.search(r'Invoice (\d+)', fname, re.IGNORECASE)
        return int(match.group(1)) if match else float('inf')

    all_invoice_pdfs = []
    processed_folders = []

    all_folders = sorted(os.listdir(source_path), key=get_folder_number)

    for folder_name in all_folders:
        folder_path = os.path.join(source_path, folder_name)
        if not os.path.isdir(folder_path): continue

        f_num = get_folder_number(folder_name)
        if f_num in valid_folders:
            found_in_folder = False
            for file_name in os.listdir(folder_path):
                if "invoice" in file_name.lower() and file_name.lower().endswith(".pdf"):
                    all_invoice_pdfs.append(os.path.join(folder_path, file_name))
                    found_in_folder = True

            if found_in_folder:
                processed_folders.append(f_num)
                print(f"Обработана папка: {folder_name}")

    if not all_invoice_pdfs:
        print("Файлы Invoice не найдены.")
        return

    # Сортировка по номеру инвойса
    all_invoice_pdfs.sort(key=lambda x: get_invoice_num(os.path.basename(x)))

    range_str = generate_range_string(processed_folders)
    output_name = f"Inv. + Spec. {range_str} {len(all_invoice_pdfs)} pcs..pdf"

    merger = PdfMerger()
    for pdf in all_invoice_pdfs:
        try:
            merger.append(pdf)
        except Exception as e:
            print(f"Ошибка добавления {pdf}: {e}")

    save_merged_pdf(merger, save_path, output_name)


# ==========================================
# ЛОГИКА 2: BindingGTDESD (Декларации и ЭСД)
# ==========================================
def process_gtd_esd(source_path, save_path, valid_folders):
    print("\n--- Запуск: Декларации и ЭСД ---")
    processed_folders = []
    all_pdfs = []

    all_folders = sorted(os.listdir(source_path), key=get_folder_number)

    for folder_name in all_folders:
        folder_path = os.path.join(source_path, folder_name)
        if not os.path.isdir(folder_path): continue

        f_num = get_folder_number(folder_name)
        if f_num in valid_folders:
            gtd_files = []
            esd_files = []

            for file_name in os.listdir(folder_path):
                if not file_name.lower().endswith(".pdf"): continue

                if file_name.startswith("GTD_"):
                    gtd_files.append(os.path.join(folder_path, file_name))
                elif file_name.count('-') == 4:  # Признак ЭСД
                    esd_files.append(os.path.join(folder_path, file_name))

            if gtd_files and esd_files:
                processed_folders.append(f_num)
                all_pdfs.extend(sorted(gtd_files)[:1])  # Берем первый GTD
                all_pdfs.extend(sorted(esd_files)[:1])  # Берем первый ESD
                print(f"Добавлена пара из папки: {folder_name}")
            else:
                print(f"Пропуск папки {folder_name}: некомплект (GTD: {len(gtd_files)}, ESD: {len(esd_files)})")

    if not all_pdfs:
        print("Не найдено пар GTD+ESD.")
        return

    range_str = generate_range_string(processed_folders)
    output_name = f"GTD+ЭСД {range_str} {len(processed_folders)} pcs..pdf"

    merger = PdfMerger()
    for pdf in all_pdfs:
        merger.append(pdf)

    save_merged_pdf(merger, save_path, output_name)


# ==========================================
# ЛОГИКА 3: BindingGTDInvSpec (GTD + Invoice + Spec)
# ==========================================
def process_gtd_inv_spec(source_path, save_path, valid_folders):
    print("\n--- Запуск: Декларации, Инвойсы и Спецификации ---")

    def get_gtd_num_from_file(fname):
        parts = re.split(r'[_-]', fname.replace('.pdf', ''))
        for part in reversed(parts):
            if part.isdigit(): return int(part)
        return float('inf')

    paired_documents = {}  # Key: GTD number
    processed_folders_set = set()

    all_folders = sorted(os.listdir(source_path), key=get_folder_number)

    for folder_name in all_folders:
        folder_path = os.path.join(source_path, folder_name)
        if not os.path.isdir(folder_path): continue

        f_num = get_folder_number(folder_name)
        if f_num in valid_folders:
            gtd_path = None
            inv_path = None

            for file_name in os.listdir(folder_path):
                lower_name = file_name.lower()
                if not lower_name.endswith(".pdf"): continue

                if lower_name.startswith("gtd_"):
                    gtd_path = os.path.join(folder_path, file_name)
                elif "invoice" in lower_name:
                    inv_path = os.path.join(folder_path, file_name)

            if gtd_path:
                gtd_num = get_gtd_num_from_file(os.path.basename(gtd_path))
                if gtd_num != float('inf'):
                    paired_documents[gtd_num] = {'gtd': gtd_path, 'invoice': inv_path}
                    processed_folders_set.add(f_num)
                    print(f"Найдено в папке {folder_name}: GTD {gtd_num} + {'Invoice' if inv_path else 'No Invoice'}")

    files_to_merge = []
    for gtd_num in sorted(paired_documents.keys()):
        pair = paired_documents[gtd_num]
        if pair['gtd']: files_to_merge.append(pair['gtd'])
        if pair['invoice']: files_to_merge.append(pair['invoice'])

    if not files_to_merge:
        print("Файлы не найдены.")
        return

    range_str = generate_range_string(list(processed_folders_set))
    output_name = f"GTD+Inv. + Spec. {range_str} {len(processed_folders_set)} pcs..pdf"

    merger = PdfMerger()
    for pdf in files_to_merge:
        merger.append(pdf)

    save_merged_pdf(merger, save_path, output_name)


# ==========================================
# ЛОГИКА 4: BindingGTD (Только Декларации)
# ==========================================
def process_gtd_only(source_path, save_path, valid_folders):
    print("\n--- Запуск: Только Декларации (GTD) ---")
    processed_folders = []
    all_pdfs = []

    all_folders = sorted(os.listdir(source_path), key=get_folder_number)

    for folder_name in all_folders:
        folder_path = os.path.join(source_path, folder_name)
        if not os.path.isdir(folder_path): continue

        f_num = get_folder_number(folder_name)
        if f_num in valid_folders:
            gtd_files = []
            for file_name in os.listdir(folder_path):
                if file_name.lower().endswith(".pdf") and file_name.startswith("GTD_"):
                    gtd_files.append(os.path.join(folder_path, file_name))

            if gtd_files:
                processed_folders.append(f_num)
                all_pdfs.extend(sorted(gtd_files)[:1])
                print(f"GTD найден в папке: {folder_name}")

    if not all_pdfs:
        print("GTD файлы не найдены.")
        return

    range_str = generate_range_string(processed_folders)
    output_name = f"GTD {range_str} {len(processed_folders)} pcs..pdf"

    merger = PdfMerger()
    for pdf in all_pdfs:
        merger.append(pdf)

    save_merged_pdf(merger, save_path, output_name)


# ==========================================
# ЛОГИКА TEMP: BindingTemp (Папка Temp)
# ==========================================
def process_temp_folder():
    print("\n--- Запуск: Скрепление из папки Temp ---")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    temp_folder = os.path.join(script_dir, "Temp")
    combined_folder = os.path.join(script_dir, "Combined")

    if not os.path.exists(temp_folder):
        print(f"❌ Ошибка: Папка Temp не найдена по пути: {temp_folder}")
        print("Создайте папку 'Temp' рядом со скриптом и положите туда файлы.")
        return

    # Логика Temp: сортировка по цифре перед запятой (1, ... или 10, ...)
    def extract_temp_number(filename):
        match = re.match(r"^(\d+),", filename)
        return int(match.group(1)) if match else float('inf')

    pdf_files = [f for f in os.listdir(temp_folder) if f.lower().endswith(".pdf")]
    sorted_pdfs = sorted(pdf_files, key=extract_temp_number)

    if not sorted_pdfs:
        print("В папке Temp нет PDF файлов.")
        return

    merger = PdfMerger()
    print("Объединяю файлы:")
    for pdf in sorted_pdfs:
        print(f"-> {pdf}")
        merger.append(os.path.join(temp_folder, pdf))

    # Генерация имени Combined-X.pdf
    if not os.path.exists(combined_folder): os.makedirs(combined_folder)

    existing = [f for f in os.listdir(combined_folder) if f.startswith("Combined") and f.endswith(".pdf")]
    next_num = 1
    if existing:
        nums = []
        for f in existing:
            m = re.search(r"Combined-(\d+)", f)
            if m: nums.append(int(m.group(1)))
        if nums: next_num = max(nums) + 1

    out_name = f"Combined-{next_num}.pdf"
    save_merged_pdf(merger, combined_folder, out_name)


# ==========================================
# ГЛАВНОЕ МЕНЮ
# ==========================================
def main():
    while True:
        print("\n" + "=" * 40)
        print("   УТИЛИТА СКРЕПЛЕНИЯ ДОКУМЕНТОВ")
        print("=" * 40)
        print("Выберите документы для скрепления:")
        print("1. Отгрузочные документы (GTD, Invoice, ESD)")
        print("2. Документы из папки Temp")
        print("0. Выход")

        choice = input("\nВаш выбор: ").strip()

        if choice == '0':
            print("Выход из программы.")
            break

        elif choice == '2':
            process_temp_folder()
            input("\nНажмите Enter, чтобы вернуться в меню...")

        elif choice == '1':
            # Шаг 1: Исходная папка
            source_path = get_clean_path("\nУкажите путь к директории с папками инвойсов")
            if not os.path.isdir(source_path):
                print("❌ Ошибка: Указанная папка не существует.")
                continue

            # Шаг 2: Диапазон
            range_input = input("Укажите диапазон номеров папок (например: 3550-3553,3560): ").strip()
            valid_folders = parse_folder_range(range_input)
            if not valid_folders:
                print("❌ Не указан корректный диапазон.")
                continue
            print(f"Будут обработаны папки: {valid_folders}")

            # Шаг 3: Выбор типа скрепления
            print("\nВыберите тип документов:")
            print("1. Инвойсы и Спецификации")
            print("2. Декларации и ЭСД")
            print("3. Декларации, Инвойсы и Спецификации")
            print("4. Декларации (Только GTD)")

            sub_choice = input("Ваш выбор: ").strip()

            # Шаг 4: Путь сохранения
            save_path = get_clean_path("Укажите путь для сохранения скрепленного файла")

            # Запуск соответствующей логики
            if sub_choice == '1':
                process_inv_spec(source_path, save_path, valid_folders)
            elif sub_choice == '2':
                process_gtd_esd(source_path, save_path, valid_folders)
            elif sub_choice == '3':
                process_gtd_inv_spec(source_path, save_path, valid_folders)
            elif sub_choice == '4':
                process_gtd_only(source_path, save_path, valid_folders)
            else:
                print("Неверный выбор типа документов.")

            input("\nГотово. Нажмите Enter, чтобы вернуться в меню...")

        else:
            print("Неверный ввод, попробуйте еще раз.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрограмма остановлена пользователем.")
