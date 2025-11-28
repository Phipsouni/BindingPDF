import os
import re
from PyPDF2 import PdfMerger

# Папка с PDF-файлами
temp_folder = os.path.join(os.getcwd(), "Temp")  # Путь к папке Temp
combined_folder = os.path.join(os.getcwd(), "Combined")  # Путь к папке Combined

# Создаём папку Combined, если её нет
os.makedirs(combined_folder, exist_ok=True)

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
existing_files = [f for f in os.listdir(combined_folder) if f.startswith("Combined") and f.endswith(".pdf")]
combined_number_pattern = re.compile(r"Combined-(\d+)\.pdf")

# Находим максимальный существующий номер
existing_numbers = [int(combined_number_pattern.search(f).group(1)) for f in existing_files if combined_number_pattern.search(f)]
next_number = max(existing_numbers, default=0) + 1

# Итоговое имя файла
output_pdf = os.path.join(combined_folder, f"Combined-{next_number}.pdf")

# Сохраняем объединённый PDF
merger.write(output_pdf)
merger.close()

print(f"\n✅ Файлы успешно объединены! Итоговый файл: {output_pdf}")
