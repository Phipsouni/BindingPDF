import os
import re
from PyPDF2 import PdfMerger

# Папка с PDF-файлами
temp_folder = os.path.join(os.getcwd(), "Temp")  # Путь к папке Temp
bound_folder = os.path.join(os.getcwd(), "Bound")  # Путь к папке Bound

# Создаём папку Bound, если её нет
os.makedirs(bound_folder, exist_ok=True)

# Регулярное выражение для извлечения числа перед запятой
number_pattern = re.compile(r"^(\d+),")

# Функция для извлечения числа перед запятой (или большого числа, если не найдено)
def extract_number(filename):
    match = number_pattern.match(filename)
    return int(match.group(1)) if match else float('inf')  # 'inf' отправляет файлы без номера в конец списка

# Получаем список всех PDF-файлов в папке Temp
pdf_files = [f for f in os.listdir(temp_folder) if f.lower().endswith(".pdf")]

# Сортируем файлы по числу перед запятой
sorted_pdfs = sorted(pdf_files, key=extract_number)

# Проверяем, есть ли PDF для объединения
if not sorted_pdfs:
    print("Ошибка: В папке 'Temp' нет PDF-файлов!")
    exit()

# Объединяем файлы в один PDF
merger = PdfMerger()

for pdf in sorted_pdfs:
    pdf_path = os.path.join(temp_folder, pdf)
    merger.append(pdf_path)
    print(f"Добавлен: {pdf}")

# --- Определяем имя выходного файла ---
existing_files = [f for f in os.listdir(bound_folder) if f.startswith("Bound") and f.endswith(".pdf")]
bound_number_pattern = re.compile(r"Bound-(\d+)\.pdf")

# Находим максимальный существующий номер
existing_numbers = [int(bound_number_pattern.search(f).group(1)) for f in existing_files if bound_number_pattern.search(f)]
next_number = max(existing_numbers, default=0) + 1

# Итоговое имя файла
output_pdf = os.path.join(bound_folder, f"Bound-{next_number}.pdf")

# Сохраняем объединённый PDF
merger.write(output_pdf)
merger.close()

print(f"\n✅ Файлы успешно объединены! Итоговый файл: {output_pdf}")