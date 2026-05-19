import openpyxl

import os
import sys
from datetime import datetime


def list_xlsx_files() -> list:
    """
    Возвращает список всех .xlsx файлов в текущей папке.
    """

    return [f for f in os.listdir('.') if f.endswith('.xlsx')]


def process_file(input_file):
    """
    Удаляет указанные столбцы и столбцы после 'Статус', сортирует строки по дате.

    Параметры:
        input_file (str): путь к исходному файлу Excel.
    """

    def get_datetime(row):
        date_val = row[date_index_in_keep]
        if date_val is None:
            return datetime.min
        # Если уже объект datetime, возвращаем как есть
        if isinstance(date_val, datetime):
            return date_val
        # Иначе пытаемся преобразовать строку
        try:
            return datetime.strptime(str(date_val), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            # на случай других форматов, например, только дата
            try:
                return datetime.strptime(str(date_val), "%Y-%m-%d")
            except ValueError:
                print(f"Не удалось распознать дату: {date_val}")
                return datetime.min

    # Определяем имя копии рабочей книги
    output_file = input_file.replace('.xlsx', '_filtered.xlsx')

    # Определяем путь сохранения копии
    output_file_path = f'./Результат/{output_file}'

    # Загружаем копию рабочей книги
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active  # берём первый лист

    # Список заголовков из первой строки
    headers = []
    for col in range(1, sheet.max_column + 1):
        val = sheet.cell(row=1, column=col).value
        headers.append(val)

    # Столбцы, которые нужно оставить (в порядке, как они идут в исходном файле)
    # Порядок следования важен для сохранения исходной последовательности колонок
    keep_columns_names = [
        "Номер звонка",
        "Звонивший",
        "Источник",
        "utm_source",
        "utm_medium",
        "utm_campaign",
        "Дата",
        "Длительность",
        "Ожидание",
        "Куда звонили",
        "Статус"
    ]

    # Находим индексы (номера столбцов) для оставляемых колонок (1-индексация)
    keep_indices = []
    for name in keep_columns_names:
        try:
            idx = headers.index(name) + 1  # +1 потому что openpyxl индексирует с 1
            keep_indices.append(idx)
        except ValueError:
            print(f"Ошибка: столбец '{name}' не найден в файле.")
            sys.exit(1)

    # Считываем все строки данных (начиная со 2-й строки)
    data_rows = []
    for row in range(2, sheet.max_row + 1):
        row_data = []
        for col_idx in keep_indices:
            cell_value = sheet.cell(row=row, column=col_idx).value
            row_data.append(cell_value)
        data_rows.append(row_data)

    # Функция для извлечения даты из строки данных (столбец "Дата" – 8-й по счёту в keep_columns_names)
    date_index_in_keep = keep_columns_names.index("Дата")  # позиция в списке keep_columns_names

    # Сортируем строки по возрастанию даты
    data_rows.sort(key=get_datetime)

    # Записываем заголовки
    for col_idx, name in enumerate(keep_columns_names, start=1):
        sheet.cell(row=1, column=col_idx, value=name)

    # Записываем данные
    for row_idx, row_data in enumerate(data_rows, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            sheet.cell(row=row_idx, column=col_idx, value=value)

    wb.save(output_file_path)
    print(f"Обработка завершена. Результат сохранён в {output_file_path}")


if __name__ == "__main__":
    call_log_files = list_xlsx_files()

    if not os.path.isdir("Результат"):
        os.mkdir("Результат")

    for filename in call_log_files:
        process_file(filename)