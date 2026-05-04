import os
import re
import sys
import time
import json  # Добавили для работы с настройками
from PyPDF2 import PdfMerger

# ==========================================
# КОНСТАНТЫ ОФОРМЛЕНИЯ И НАСТРОЕК
# ==========================================
BOLD = "\033[1m"
RESET = "\033[0m"

# Определяем путь к папке, где лежит сам скрипт
script_dir = os.path.dirname(os.path.abspath(__file__))
# Полный путь к конфигу (чтобы он создавался рядом со скриптом)
CONFIG_FILE = os.path.join(script_dir, "config.json")

# ==========================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (Утилиты)
# ==========================================

def print_error(message):
    """Вывод ошибки с красным восклицательным знаком."""
    print(f"❗️ {message}")


def load_config():
    """Загружает настройки из JSON файла."""
    if not os.path.exists(CONFIG_FILE):
        return None
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print_error(f"Ошибка чтения конфига: {e}")
        return None


def save_config(source_path, save_path):
    """Сохраняет пути в JSON файл."""
    data = {
        "source_path": source_path,
        "save_path": save_path
    }
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        # print(f"{BOLD}✔ Настройки путей сохранены.{RESET}") # Можно раскомментировать для отладки
    except Exception as e:
        print_error(f"Не удалось сохранить настройки: {e}")


def get_clean_path(prompt_text, allow_menu_codes=False):
    """
    Запрашивает путь.
    Текст запроса (prompt_text) делается жирным.
    """
    full_prompt = f"{BOLD}{prompt_text}:{RESET} "
    path = input(full_prompt).strip()

    if allow_menu_codes and path in ['0', '1', '9']:  # Добавили 9 для сброса
        return path

    if (path.startswith('"') and path.endswith('"')) or (path.startswith("'") and path.endswith("'")):
        path = path[1:-1]
    return path


def parse_folder_range(range_str):
    """Парсит строку диапазона (например, '3550-3553,3560')."""
    ranges = range_str.split(',')
    folder_numbers = set()
    for r in ranges:
        r = r.strip()
        if not r: continue
        if '-' in r:
            try:
                parts = r.split('-')
                start, end = int(parts[0]), int(parts[1])
                if start > end:
                    print_error(f"Неверный диапазон '{r}'.")
                    continue
                folder_numbers.update(range(start, end + 1))
            except ValueError:
                print_error(f"Неверный формат диапазона '{r}'.")
        else:
            try:
                folder_numbers.add(int(r))
            except ValueError:
                print_error(f"Неверный формат числа '{r}'.")
    return sorted(list(folder_numbers))


def get_number_from_string(text):
    """Извлекает первое число из строки для сортировки."""
    match = re.search(r'\d+', text)
    return int(match.group()) if match else float('inf')


def generate_range_string(processed_numbers):
    """Генерирует строку диапазона (3550-3552;3560)."""
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

        print(f"Сохранение: {file_name} ...")
        with open(full_path, 'wb') as f_out:
            merger.write(f_out)
        merger.close()
        print(f"✅ Готово!")
    except Exception as e:
        print_error(f"Ошибка при сохранении: {e}")


# ==========================================
# ЛОГИКА 1: BindingInvSpec (Инвойсы и Спецификации)
# ==========================================
def process_inv_spec(source_path, save_path, valid_folders):
    print("\n[Выполняется: Инвойсы и Спецификации]")

    def get_invoice_num(fname):
        match = re.search(r'Invoice (\d+)', fname, re.IGNORECASE)
        return int(match.group(1)) if match else float('inf')

    all_invoice_pdfs = []
    processed_folders = []

    all_folders = sorted(os.listdir(source_path), key=get_number_from_string)

    for folder_name in all_folders:
        folder_path = os.path.join(source_path, folder_name)
        if not os.path.isdir(folder_path): continue

        f_num = get_number_from_string(folder_name)
        if f_num in valid_folders:
            found_in_folder = False
            for file_name in os.listdir(folder_path):
                if "invoice" in file_name.lower() and file_name.lower().endswith(".pdf"):
                    all_invoice_pdfs.append(os.path.join(folder_path, file_name))
                    found_in_folder = True

            if found_in_folder:
                processed_folders.append(f_num)

    if not all_invoice_pdfs:
        print_error("Файлы Invoice не найдены.")
        return

    all_invoice_pdfs.sort(key=lambda x: get_invoice_num(os.path.basename(x)))

    range_str = generate_range_string(processed_folders)
    output_name = f"Inv. + Spec. {range_str} {len(all_invoice_pdfs)} pcs..pdf"

    merger = PdfMerger()
    for pdf in all_invoice_pdfs:
        try:
            merger.append(pdf)
        except Exception as e:
            print_error(f"Ошибка с файлом {pdf}: {e}")

    save_merged_pdf(merger, save_path, output_name)


# ==========================================
# ЛОГИКА 2: BindingGTDESD (Декларации и ЭСД)
# ==========================================
def process_gtd_esd(source_path, save_path, valid_folders):
    print("\n[Выполняется: Декларации и ЭСД]")
    processed_folders = []
    all_pdfs = []

    all_folders = sorted(os.listdir(source_path), key=get_number_from_string)

    for folder_name in all_folders:
        folder_path = os.path.join(source_path, folder_name)
        if not os.path.isdir(folder_path): continue

        f_num = get_number_from_string(folder_name)
        if f_num in valid_folders:
            gtd_files = []
            esd_files = []

            for file_name in os.listdir(folder_path):
                if not file_name.lower().endswith(".pdf"): continue
                if file_name.startswith("GTD_"):
                    gtd_files.append(os.path.join(folder_path, file_name))
                elif file_name.count('-') == 4:
                    esd_files.append(os.path.join(folder_path, file_name))

            if gtd_files and esd_files:
                processed_folders.append(f_num)
                all_pdfs.extend(sorted(gtd_files)[:1])
                all_pdfs.extend(sorted(esd_files)[:1])
            else:
                print_error(f"Папка {folder_name} пропущена: некомплект.")

    if not all_pdfs:
        print_error("Не найдено пар GTD+ESD.")
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
    print("\n[Выполняется: Декларации, Инвойсы и Спецификации]")

    def parse_gtd_date_ddmmyy(s):
        """Шестизначная дата ДДММГГ в имени GTD → сортируемый кортеж (год, месяц, день)."""
        if len(s) != 6 or not s.isdigit():
            return None
        d, m, y = int(s[:2]), int(s[2:4]), int(s[4:6])
        if not (1 <= m <= 12 and 1 <= d <= 31):
            return None
        return (y, m, d)

    def get_gtd_sort_key_from_file(fname):
        """
        Ключ сортировки для GTD: сначала дата (средний блок ДДММГГ),
        затем код таможни (первый числовой блок), затем порядковый номер (последний блок).
        Пример: GTD_10137010_300426_5011010 → дата 30.04.26, таможня 10137010, номер 5011010.
        Если формат не распознан — как раньше, только по последнему числу в имени.
        """
        base = os.path.basename(fname)
        if base.lower().endswith(".pdf"):
            base = base[:-4]
        parts = re.split(r"[_-]", base)
        numeric_parts = [part for part in parts if part.isdigit()]
        if len(parts) >= 2 and parts[-1].isdigit() and parts[-2].isdigit():
            reg = int(parts[-1])
            date_tuple = parse_gtd_date_ddmmyy(parts[-2])
            if date_tuple is not None:
                customs_code = int(numeric_parts[0]) if len(numeric_parts) >= 3 else float("inf")
                return (0, date_tuple, customs_code, reg)
        for part in reversed(parts):
            if part.isdigit():
                return (1, (9999, 99, 99), float("inf"), int(part))
        return (1, (9999, 99, 99), float("inf"), float("inf"))

    valid_pairs = []
    processed_folders_set = set()

    all_folders = sorted(os.listdir(source_path), key=get_number_from_string)

    for folder_name in all_folders:
        folder_path = os.path.join(source_path, folder_name)
        if not os.path.isdir(folder_path): continue

        f_num = get_number_from_string(folder_name)
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

            if gtd_path and inv_path:
                sort_key = get_gtd_sort_key_from_file(os.path.basename(gtd_path))
                valid_pairs.append({
                    'sort_key': sort_key,
                    'gtd': gtd_path,
                    'inv': inv_path
                })
                processed_folders_set.add(f_num)

            elif gtd_path and not inv_path:
                print_error(f"Папка {folder_name}: Найден GTD, но нет Invoice! (Пропущено)")

            elif inv_path and not gtd_path:
                print_error(f"Папка {folder_name}: Найден Invoice, но нет GTD! (Пропущено)")

    if not valid_pairs:
        print_error("Файлы для скрепления не найдены (или возникли ошибки комплектности).")
        return

    valid_pairs.sort(key=lambda x: x["sort_key"])

    files_to_merge = []
    for pair in valid_pairs:
        files_to_merge.append(pair['gtd'])
        files_to_merge.append(pair['inv'])

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
    print("\n[Выполняется: Только Декларации (GTD)]")
    processed_folders = []
    all_pdfs = []

    all_folders = sorted(os.listdir(source_path), key=get_number_from_string)

    for folder_name in all_folders:
        folder_path = os.path.join(source_path, folder_name)
        if not os.path.isdir(folder_path): continue

        f_num = get_number_from_string(folder_name)
        if f_num in valid_folders:
            gtd_files = []
            for file_name in os.listdir(folder_path):
                if file_name.lower().endswith(".pdf") and file_name.startswith("GTD_"):
                    gtd_files.append(os.path.join(folder_path, file_name))

            if gtd_files:
                processed_folders.append(f_num)
                all_pdfs.extend(sorted(gtd_files)[:1])

    if not all_pdfs:
        print_error("GTD файлы не найдены.")
        return

    range_str = generate_range_string(processed_folders)
    output_name = f"GTD {range_str} {len(processed_folders)} pcs..pdf"

    merger = PdfMerger()
    for pdf in all_pdfs:
        merger.append(pdf)

    save_merged_pdf(merger, save_path, output_name)


# ==========================================
# ЛОГИКА 5: Ж/Д Накладные (Railway)
# ==========================================
def process_railway():
    print("\n[Выполняется: Ж/Д накладные по 4 шт.]")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    source_folder = os.path.join(script_dir, "Railway")
    save_folder = os.path.join(script_dir, "Merged Railway")

    if not os.path.exists(source_folder):
        print_error(f"Папка Railway не найдена по пути: {source_folder}")
        print("Создайте папку 'Railway' рядом со скриптом.")
        return

    files = [f for f in os.listdir(source_folder) if f.lower().endswith('.pdf')]
    if not files:
        print_error("В папке Railway нет PDF файлов.")
        return

    files.sort(key=get_number_from_string)

    chunk_size = 4
    chunks = [files[i:i + chunk_size] for i in range(0, len(files), chunk_size)]

    if not os.path.exists(save_folder):
        os.makedirs(save_folder)

    for chunk in chunks:
        merger = PdfMerger()
        file_numbers = []

        print(f"Скрепляю ({len(chunk)} шт): {chunk[0]} ... {chunk[-1]}")

        for fname in chunk:
            full_path = os.path.join(source_folder, fname)
            merger.append(full_path)
            file_numbers.append(get_number_from_string(fname))

        range_str = generate_range_string(file_numbers)
        output_name = f"Railway {range_str} {len(chunk)} pcs..pdf"

        save_merged_pdf(merger, save_folder, output_name)

    print(f"\n✅ Все файлы обработаны. Сохранено в: {save_folder}")


# ==========================================
# ЛОГИКА TEMP (Папка Temp)
# ==========================================
def process_temp_folder():
    print("\n[Выполняется: Скрепление из папки Temp]")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    temp_folder = os.path.join(script_dir, "Temp")
    combined_folder = os.path.join(script_dir, "Combined")

    if not os.path.exists(temp_folder):
        print_error("Папка Temp не найдена.")
        return

    def extract_temp_number(filename):
        match = re.match(r"^(\d+),", filename)
        return int(match.group(1)) if match else float('inf')

    pdf_files = [f for f in os.listdir(temp_folder) if f.lower().endswith(".pdf")]
    sorted_pdfs = sorted(pdf_files, key=extract_temp_number)

    if not sorted_pdfs:
        print_error("В папке Temp нет PDF файлов.")
        return

    merger = PdfMerger()
    for pdf in sorted_pdfs:
        merger.append(os.path.join(temp_folder, pdf))

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
# МЕНЮ (STATE MACHINE)
# ==========================================

def main():
    while True:
        # --- ГЛАВНОЕ МЕНЮ ---
        print("\n" + "=" * 45)
        print(f"   {BOLD}УТИЛИТА СКРЕПЛЕНИЯ ДОКУМЕНТОВ{RESET}")
        print("=" * 45)
        print("Выберите документы для скрепления:")
        print("1. Отгрузочные документы (GTD, Invoice, ESD)")
        print("2. Документы из папки Temp")
        print("3. Ж/Д накладные из папки Railway")
        print("0. Выход")

        main_choice = input(f"\n{BOLD}Ваш выбор:{RESET} ").strip()

        if main_choice == '0':
            print("Выход из программы.")
            break

        elif main_choice == '2':
            process_temp_folder()

        elif main_choice == '3':
            process_railway()

        elif main_choice == '1':
            shipping_docs_workflow()

        else:
            print_error("Неверный ввод.")


def shipping_docs_workflow():
    """
    Машина состояний для процесса отгрузочных документов.
    """

    # ----------------------------------------
    # ИНИЦИАЛИЗАЦИЯ: Проверка сохраненных путей
    # ----------------------------------------
    source_path = ""
    save_path = ""
    valid_folders = []

    config = load_config()
    loaded_from_config = False

    if config:
        cfg_source = config.get("source_path", "")
        cfg_save = config.get("save_path", "")

        # Проверяем, существуют ли сохраненные папки
        if os.path.isdir(cfg_source) and os.path.isdir(cfg_save):
            source_path = cfg_source
            save_path = cfg_save
            loaded_from_config = True
            current_state = 'ASK_RANGE'  # Сразу переходим к диапазону
            print(f"\nℹ️  {BOLD}Используются сохраненные пути:{RESET}")
            print(f"   📁 Откуда: {source_path}")
            print(f"   💾 Куда:   {save_path}")
        else:
            current_state = 'ASK_SOURCE'
    else:
        current_state = 'ASK_SOURCE'

    while True:
        # ----------------------------------------
        # ЭТАП 1: Выбор исходной папки
        # ----------------------------------------
        if current_state == 'ASK_SOURCE':
            print("\n🟧 Шаг 1: Исходная директория")
            print("Введите путь или '0' для возврата в Главное Меню.")
            user_input = get_clean_path("Путь к папке с инвойсами", allow_menu_codes=True)

            if user_input == '0':
                return  # Возврат в main()

            if not os.path.isdir(user_input):
                print_error("Папка не существует. Попробуйте еще раз.")
                continue

            source_path = user_input
            current_state = 'ASK_SAVE_PATH'

        # ----------------------------------------
        # ЭТАП 2: Путь сохранения
        # ----------------------------------------
        elif current_state == 'ASK_SAVE_PATH':
            print("\n🟧 Шаг 2: Место сохранения")
            print("Введите путь для сохранения готового файла.")
            print("1. Возврат к выбору исходной директории")
            print("0. Возврат в главное меню")

            user_input = get_clean_path("Путь сохранения", allow_menu_codes=True)

            if user_input == '0': return
            if user_input == '1':
                current_state = 'ASK_SOURCE'
                continue

            save_path = user_input

            # Если оба пути валидны, сохраняем конфиг
            if os.path.isdir(source_path) and os.path.isdir(save_path):
                save_config(source_path, save_path)

            current_state = 'ASK_RANGE'

        # ----------------------------------------
        # ЭТАП 3: Выбор диапазона
        # ----------------------------------------
        elif current_state == 'ASK_RANGE':
            print("\n🟧 Шаг 3: Диапазон папок")
            print("Введите диапазон (например: 3550-3553,3560)")
            print("1. Изменить путь сохранения (Назад)")
            print("9. Сбросить все пути и выбрать папку заново")
            print("0. Возврат в главное меню")

            user_input = input(f"{BOLD}Диапазон или команда:{RESET} ").strip()

            if user_input == '0': return
            if user_input == '1':
                current_state = 'ASK_SAVE_PATH'
                continue
            if user_input == '9':
                current_state = 'ASK_SOURCE'
                continue

            folders = parse_folder_range(user_input)
            if not folders:
                print_error("Некорректный диапазон.")
                continue

            print(f"✔ Будут обработаны папки: {folders}")
            valid_folders = folders
            current_state = 'SELECT_TYPE'

        # ----------------------------------------
        # ЭТАП 4: Выбор типа скрепления
        # ----------------------------------------
        elif current_state == 'SELECT_TYPE':
            print("\n🟧 Шаг 4: Выбор типа скрепления")

            # Показываем короткие версии путей для наглядности
            short_source = source_path[-30:] if len(source_path) > 30 else source_path
            short_save = save_path[-30:] if len(save_path) > 30 else save_path

            print(f"Источник: ...{short_source}")
            print(f"Сохранение в: ...{short_save}")
            print("-" * 30)
            print("1. Инвойсы и Спецификации")
            print("2. Декларации и ЭСД")
            print("3. Декларации, Инвойсы и Спецификации")
            print("4. Декларации (Только GTD)")
            print("-" * 30)
            print("6. Возврат к выбору диапазона номеров")
            print("7. Изменить пути (возврат к выбору папки)")
            print("0. Возврат в главное меню")

            choice = input(f"\n{BOLD}Ваш выбор:{RESET} ").strip()

            # Навигация
            if choice == '0': return
            if choice == '6':
                current_state = 'ASK_RANGE'
                continue
            if choice == '7':
                current_state = 'ASK_SOURCE'
                continue

            # Действия
            if choice == '1':
                process_inv_spec(source_path, save_path, valid_folders)
            elif choice == '2':
                process_gtd_esd(source_path, save_path, valid_folders)
            elif choice == '3':
                process_gtd_inv_spec(source_path, save_path, valid_folders)
            elif choice == '4':
                process_gtd_only(source_path, save_path, valid_folders)
            else:
                print_error("Неверный выбор.")
                time.sleep(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрограмма остановлена.")
