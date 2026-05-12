import requests
import pdfplumber
import pandas as pd
from io import BytesIO
import urllib3
import re

# Отключаем предупреждения о SSL (если сайт использует самоподписанный сертификат)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def clean_numbers_from_line(line: str):
    """
    Извлекает все числа из строки, нормализуя пробелы внутри чисел.
    Например: "5 000" -> "5000"
    Возвращает список целых чисел.
    """
    # Сначала убираем пробелы между цифрами (5 000 -> 5000, 12 000 -> 12000)
    line_no_spaces_in_numbers = re.sub(r'(\d)\s+(?=\d)', r'\1', line)
    # Находим все последовательности цифр
    numbers = re.findall(r'\d+', line_no_spaces_in_numbers)
    # Преобразуем в int
    return [int(num) for num in numbers]

def parse_data_line(line: str):
    """
    Парсит одну строку таблицы.
    Возвращает словарь с полями: тип, название, часы (строка), цена (строка).
    Если строка не содержит данных (нет пар часы+цена) – возвращает None.
    """
    line = line.strip()
    if not line:
        return None

    # Извлекаем все числа из строки
    numbers = clean_numbers_from_line(line)
    if not numbers:
        return None

    # Разбиваем числа на пары (часы, цена)
    pairs = []
    for i in range(0, len(numbers) - 1, 2):
        hours = numbers[i]
        price = numbers[i + 1]
        pairs.append((hours, price))

    if not pairs:
        return None

    # Формируем строки с часами и ценами, объединяя пары через разделитель " / "
    hours_str = ' / '.join(str(h) for h, _ in pairs)
    price_str = ' / '.join(str(p) for _, p in pairs)

    # Теперь нужно выделить тип программы и название.
    # Для этого убираем из строки все найденные числа (они уже извлечены).
    # Сначала убираем пробелы между цифрами, как и ранее
    line_clean = re.sub(r'(\d)\s+(?=\d)', r'\1', line)
    # Удаляем все цифры и лишние пробелы
    without_numbers = re.sub(r'\d+', '', line_clean)
    without_numbers = re.sub(r'\s+', ' ', without_numbers).strip()

    # В начале строки может быть указание типа: например "ОР*", "ПК/", "ПП*"
    # Пытаемся найти тип как последовательность из 2-4 букв (русских или латинских) и, возможно, символов * /
    type_match = re.match(r'^([A-Za-zА-Яа-я]{2,4}[\*/]?)\s+', without_numbers)
    prog_type = ''
    name_part = without_numbers
    if type_match:
        prog_type = type_match.group(1)
        name_part = without_numbers[type_match.end():].strip()

    return {
        'Тип программы': prog_type,
        'Наименование программы': name_part,
        'Кол-во часов': hours_str,
        'Стоимость, руб.': price_str
    }

def main():
    url = "https://anodpo.ru/price.pdf"
    print(f"Загрузка PDF: {url}")
    response = requests.get(url, verify=False, timeout=30)
    response.raise_for_status()
    print("PDF успешно загружен, начинаю парсинг...")

    all_data = []

    with pdfplumber.open(BytesIO(response.content)) as pdf:
        # Данные начинаются с 4-й страницы (индекс 3)
        for page_num in range(3, len(pdf.pages)):
            page = pdf.pages[page_num]
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            for line in lines:
                parsed = parse_data_line(line)
                if parsed:
                    parsed['Страница'] = page_num + 1
                    all_data.append(parsed)

    if not all_data:
        print("Не удалось извлечь данные. Проверьте структуру PDF.")
        return

    # Создаём DataFrame
    df = pd.DataFrame(all_data)

    # Сохраняем в Excel
    output_file = "anodpo_prices.xlsx"
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"✅ Готово! Сохранено {len(df)} записей в файл '{output_file}'")

if __name__ == "__main__":
    main()