import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

# Соответствие URL -> название категории
URL_TO_CATEGORY = {
    "elektro.html": "Электротехническое направление",
    "gpm.html": "Грузоподъемные машины и механизмы",
    "teplo.html": "Теплотехническое направление",
    "e-audit.html": "Техника и технология наземного транспорта",
    "kcn.html": "Охрана труда и пожарная безопасность",
    "business.html": "Бизнес образование",
    "prombez.html": "Промышленная безопасность",
}

PAGES_TO_PARSE = [f"https://energetik-uc.narod.ru/obuchen/{page}" for page in URL_TO_CATEGORY.keys()]

def get_soup(url):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()
    try:
        html = resp.content.decode('windows-1251')
    except UnicodeDecodeError:
        html = resp.content.decode('utf-8', errors='replace')
    return BeautifulSoup(html, 'html.parser')

def extract_number(text):
    if not text:
        return None
    match = re.search(r'\d+', text)
    return int(match.group()) if match else None

def get_category_from_row(row):
    """Определяет строку-заголовок подкатегории (внутри таблицы)"""
    cells = row.find_all('td')
    if not cells:
        return None
    first_cell = cells[0]
    colspan = int(first_cell.get('colspan', 1))
    if colspan >= 4:
        cat_text = first_cell.get_text(strip=True)
        if cat_text and any(kw in cat_text.lower() for kw in 
                            ['профессиональное обучение', 'дополнительное профессиональное образование',
                             'дополнительное образование', 'повышение квалификации', 'профессиональная подготовка']):
            return cat_text
    return None

def parse_educational_page(url, main_category):
    """
    Парсит страницу, возвращает список программ с едиными полями.
    main_category - категория, заданная по URL.
    """
    print(f"Парсинг: {url} -> {main_category}")
    soup = get_soup(url)
    if not soup:
        return []

    tables = soup.find_all('table')
    all_programs = []
    current_subcategory = ""

    for table in tables:
        rows = table.find_all('tr')
        if len(rows) < 3:
            continue

        # Поиск строки с заголовками таблицы
        header_row_idx = -1
        for idx, row in enumerate(rows[:4]):
            if any('Наименование' in cell.get_text(strip=True) for cell in row.find_all('td')):
                header_row_idx = idx
                break
        if header_row_idx == -1:
            continue

        # Заголовки (оригинальные имена, затем преобразуем в унифицированные)
        raw_headers = []
        header_cells = rows[header_row_idx].find_all('td')
        for cell in header_cells:
            colspan = int(cell.get('colspan', 1))
            text = cell.get_text(strip=True).lower()
            if 'наименование' in text:
                raw_headers.extend(['title'] * colspan)
            elif 'код' in text:
                raw_headers.extend(['code'] * colspan)
            elif 'вид подготовки' in text or 'вид' in text:
                raw_headers.extend(['training_type'] * colspan)
            elif 'дней' in text:
                raw_headers.extend(['days'] * colspan)
            elif 'часов' in text and 'всего' not in text:
                raw_headers.extend(['hours'] * colspan)
            elif 'всего' in text and 'часов' in text:
                raw_headers.extend(['total_hours'] * colspan)
            else:
                raw_headers.extend([f'extra_{len(raw_headers)}'] * colspan)
        while len(raw_headers) < 6:
            raw_headers.append(f'extra_{len(raw_headers)}')

        # Обработка строк таблицы
        active_rowspans = {}
        for row in rows[header_row_idx+1:]:
            # Проверка на строку-подкатегорию (внутри таблицы)
            subcat = get_category_from_row(row)
            if subcat:
                current_subcategory = subcat
                continue

            cells = row.find_all('td')
            if not cells:
                continue

            row_data = []
            col_index = 0
            while col_index in active_rowspans:
                row_data.append(active_rowspans[col_index]['text'])
                active_rowspans[col_index]['remaining'] -= 1
                if active_rowspans[col_index]['remaining'] == 0:
                    del active_rowspans[col_index]
                col_index += 1

            for cell in cells:
                while col_index in active_rowspans:
                    row_data.append(active_rowspans[col_index]['text'])
                    active_rowspans[col_index]['remaining'] -= 1
                    if active_rowspans[col_index]['remaining'] == 0:
                        del active_rowspans[col_index]
                    col_index += 1

                cell_text = cell.get_text(strip=True)
                rowspan = int(cell.get('rowspan', 1))
                if rowspan > 1:
                    active_rowspans[col_index] = {'text': cell_text, 'remaining': rowspan - 1}
                row_data.append(cell_text)
                col_index += 1

            while col_index in active_rowspans:
                row_data.append(active_rowspans[col_index]['text'])
                active_rowspans[col_index]['remaining'] -= 1
                if active_rowspans[col_index]['remaining'] == 0:
                    del active_rowspans[col_index]
                col_index += 1

            while len(row_data) < len(raw_headers):
                row_data.append('')

            # Создаём словарь с оригинальными ключами
            program_raw = {raw_headers[i]: row_data[i].strip() for i in range(len(raw_headers))}
            title = program_raw.get('title', '')
            if not title or len(title) < 3:
                continue

            # Унифицируем
            unified = {
                'Наименование программы': title,
                'Код профессии': program_raw.get('code', ''),
                'Вид подготовки': program_raw.get('training_type', ''),
                'Категория': main_category,
                'Подкатегория': current_subcategory,
                'Дни (текст)': program_raw.get('days', ''),
                'Часы (текст)': program_raw.get('hours', ''),
                'Всего часов (текст)': program_raw.get('total_hours', ''),
                'Источник': url
            }
            # Числовые версии
            unified['Дни (число)'] = extract_number(unified['Дни (текст)'])
            unified['Часы (число)'] = extract_number(unified['Часы (текст)'])
            unified['Всего часов (число)'] = extract_number(unified['Всего часов (текст)'])

            all_programs.append(unified)

    return all_programs

def main():
    all_records = []
    for url in PAGES_TO_PARSE:
        # Извлекаем имя файла из URL
        page_filename = url.split('/')[-1]
        main_category = URL_TO_CATEGORY.get(page_filename, "Другое")
        try:
            records = parse_educational_page(url, main_category)
            all_records.extend(records)
            print(f"  -> Найдено программ: {len(records)}")
        except Exception as e:
            print(f"Ошибка на {url}: {e}")

    if not all_records:
        print("Нет данных для сохранения.")
        return

    df = pd.DataFrame(all_records)
    # Порядок колонок
    column_order = [
        'Наименование программы', 'Код профессии', 'Вид подготовки',
        'Категория', 'Подкатегория',
        'Дни (текст)', 'Часы (текст)', 'Всего часов (текст)',
        'Дни (число)', 'Часы (число)', 'Всего часов (число)',
        'Источник'
    ]
    df = df[column_order]

    # Сохраняем в Excel (один лист)
    with pd.ExcelWriter('educational_programs.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Все программы', index=False)
        # Автоширина колонок
        worksheet = writer.sheets['Все программы']
        for column in worksheet.columns:
            max_len = 0
            col_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_len:
                        max_len = len(str(cell.value))
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = min(max_len + 2, 60)

    print(f"Сохранено educational_programs.xlsx ({len(df)} записей)")
    print("Готово!")

if __name__ == "__main__":
    main()