import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from openpyxl import Workbook


def fetch_page_content(url):
    """Загружает HTML-код страницы."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        response.encoding = 'utf-8'
        return response.text
    except requests.RequestException as e:
        print(f"Ошибка при загрузке страницы: {e}")
        return None


def parse_tables_from_html(html_content):
    """
    Извлекает таблицы из HTML-кода.
    Ищет все элементы <table> и преобразует их в DataFrame.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    tables = soup.find_all('table')
    dataframes = []

    for i, table in enumerate(tables):
        try:
            # Извлечение данных из таблицы
            rows = []
            for row in table.find_all('tr'):
                cells = row.find_all(['td', 'th'])
                row_data = [cell.get_text(strip=True) for cell in cells]
                if row_data:  # Пропускаем пустые строки
                    rows.append(row_data)

            if rows:
                # Первая строка - заголовки
                df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame(rows)
                dataframes.append((f"Table_{i+1}", df))
        except Exception as e:
            print(f"Ошибка при разборе таблицы {i+1}: {e}")

    return dataframes


def parse_text_based_tables(html_content):
    """
    Анализирует текстовое содержимое страницы для поиска разделов.
    Разделы определяются по заголовкам, написанным ЗАГЛАВНЫМИ БУКВАМИ.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    text = soup.get_text(separator='\n')

    # Регулярное выражение для поиска заголовков разделов
    # Ищем строки, состоящие из заглавных букв, цифр, пробелов и знаков препинания
    # Исключаем строки, которые могут быть частью программы
    header_pattern = r'^[А-ЯЁA-Z0-9\s\-–—.,;:()!?]+$'

    lines = text.split('\n')
    sections = []
    current_section = []
    current_header = None

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Проверяем, является ли строка заголовком раздела
        if re.match(header_pattern, line) and len(line.split()) > 1 and len(line) > 5:
            # Сохраняем предыдущий раздел, если он есть
            if current_header and current_section:
                sections.append((current_header, '\n'.join(current_section)))
            # Начинаем новый раздел
            current_header = line
            current_section = []
        else:
            if current_header is not None:
                current_section.append(line)
            else:
                # Если заголовка еще нет, пропускаем строки до первого заголовка
                continue

    # Добавляем последний раздел
    if current_header and current_section:
        sections.append((current_header, '\n'.join(current_section)))

    # Преобразуем текстовые разделы в DataFrame
    dataframes = []
    for header, content in sections:
        try:
            # Разбиваем содержимое на строки
            lines = content.split('\n')
            table_data = []

            for line in lines:
                # Пытаемся разделить строку по пробелам или табуляции
                # Учитываем, что данные могут содержать пробелы
                parts = re.split(r'\s{2,}|\t', line)
                if len(parts) >= 3:  # Минимальное количество столбцов
                    table_data.append(parts)

            if table_data:
                # Определяем максимальное количество столбцов
                max_cols = max(len(row) for row in table_data)
                # Выравниваем строки
                for row in table_data:
                    while len(row) < max_cols:
                        row.append('')

                # Создаем DataFrame
                df = pd.DataFrame(table_data)
                # Используем первую строку как заголовки, если это целесообразно
                if len(df) > 1 and all(isinstance(cell, str) for cell in df.iloc[0]):
                    df.columns = df.iloc[0]
                    df = df[1:]

                dataframes.append((header, df))
        except Exception as e:
            print(f"Ошибка при разборе раздела '{header}': {e}")

    return dataframes


def save_to_excel(dataframes, filename="educational_programs.xlsx"):
    """Сохраняет список DataFrame в Excel-файл, каждый в отдельную вкладку."""
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for sheet_name, df in dataframes:
            # Очищаем имя листа от недопустимых символов
            clean_name = re.sub(r'[\\/*?:\[\]]', '', sheet_name)[:31]  # Максимум 31 символ
            df.to_excel(writer, sheet_name=clean_name, index=False)
    print(f"Данные сохранены в файл: {filename}")


def main():
    url = "https://stc-patriot.ru/prajs-list/"
    print(f"Загрузка страницы: {url}")

    html_content = fetch_page_content(url)
    if not html_content:
        return

    print("Поиск HTML-таблиц...")
    html_tables = parse_tables_from_html(html_content)

    print("Поиск текстовых разделов...")
    text_sections = parse_text_based_tables(html_content)

    # Объединяем результаты
    all_tables = html_tables + text_sections

    if not all_tables:
        print("Таблицы не найдены.")
        return

    print(f"Найдено {len(all_tables)} таблиц/разделов.")
    save_to_excel(all_tables)


if __name__ == "__main__":
    main()