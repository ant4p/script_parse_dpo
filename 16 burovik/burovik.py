import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import urllib3
from urllib3.exceptions import InsecureRequestWarning

urllib3.disable_warnings(InsecureRequestWarning)

URL = "https://уц-буровик.рф/курсы/"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def extract_hours(duration_text):
    """
    Извлекает количество часов для подготовки, переподготовки и повышения квалификации
    из строки вида 'подготовка 290 часов, переподготовка 260 часов, повышение квалификации 80 часов'
    Возвращает кортеж (preparation, retraining, advanced)
    """
    preparation = retraining = advanced = ''
    
    # Поиск с учётом разных падежей: "час", "часа", "часов"
    prep_match = re.search(r'подготовка\s+(\d+)\s*час', duration_text, re.IGNORECASE)
    if prep_match:
        preparation = prep_match.group(1)
    
    retrain_match = re.search(r'переподготовка\s+(\d+)\s*час', duration_text, re.IGNORECASE)
    if retrain_match:
        retraining = retrain_match.group(1)
    
    adv_match = re.search(r'повышение квалификации\s+(\d+)\s*час', duration_text, re.IGNORECASE)
    if adv_match:
        advanced = adv_match.group(1)
    
    return preparation, retraining, advanced

def parse_courses():
    try:
        response = requests.get(URL, headers=HEADERS, timeout=15, verify=False)
        response.raise_for_status()
        response.encoding = 'utf-8'
    except Exception as e:
        print(f"Ошибка загрузки: {e}")
        return []

    soup = BeautifulSoup(response.text, 'html.parser')
    spoilers = soup.find_all('div', class_='su-spoiler-content')
    if not spoilers:
        print("Не найдено блоков .su-spoiler-content")
        return []

    courses_data = []

    for spoiler in spoilers:
        items = spoiler.find_all('li', recursive=True)
        for li in items:
            # Название программы
            title_tag = li.find('strong')
            if title_tag:
                course_name = title_tag.get_text(strip=True)
            else:
                course_name = li.get_text(strip=True).split('Форма обучения')[0].strip()

            # Поиск таблицы
            table = li.find('div', class_='su-table')
            if not table:
                table = li.find('table')
            if not table:
                print(f"Пропуск: {course_name} — нет таблицы")
                continue

            rows = table.find_all('tr')
            if not rows:
                continue

            # Определяем формат таблицы
            first_row_cells = rows[0].find_all('td')
            if len(first_row_cells) == 3 and \
               any('Форма' in cell.get_text() for cell in first_row_cells) and \
               any('Срок' in cell.get_text() for cell in first_row_cells) and \
               any('Стоимость' in cell.get_text() for cell in first_row_cells):
                # Формат 2 (заголовки в первой строке)
                if len(rows) >= 2:
                    value_cells = rows[1].find_all('td')
                    if len(value_cells) == 3:
                        form = value_cells[0].get_text(strip=True)
                        duration = value_cells[1].get_text(strip=True)
                        cost = value_cells[2].get_text(strip=True)
                    else:
                        form = duration = cost = ''
                else:
                    form = duration = cost = ''
            else:
                # Формат 1 (пары метка-значение)
                form = duration = cost = ''
                for row in rows:
                    cells = row.find_all('td')
                    if len(cells) < 2:
                        continue
                    label = cells[0].get_text(strip=True)
                    value = cells[1].get_text(strip=True)
                    if 'Форма обучения' in label:
                        form = value
                    elif 'Срок обучения' in label:
                        duration = value
                    elif 'Стоимость обучения' in label:
                        cost = value

            # Извлекаем часы из duration
            prep, retrain, advanced = extract_hours(duration)

            courses_data.append({
                'Наименование программы': course_name,
                'Форма обучения': form,
                'Срок обучения': duration,
                'Подготовка (часы)': prep,
                'Переподготовка (часы)': retrain,
                'Повышение квалификации (часы)': advanced,
                'Стоимость обучения': cost
            })

    return courses_data

def save_to_excel(data, filename='образовательные_программы.xlsx'):
    if not data:
        print("Нет данных для сохранения")
        return
    df = pd.DataFrame(data)
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Курсы')
        worksheet = writer.sheets['Курсы']
        for column in worksheet.columns:
            max_len = 0
            col_letter = column[0].column_letter
            for cell in column:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = min(max_len + 2, 50)
    print(f"Сохранено {len(data)} записей в {filename}")

if __name__ == '__main__':
    courses = parse_courses()
    if courses:
        print(f"Найдено программ: {len(courses)}")
        for c in courses[:3]:
            print(c)
        save_to_excel(courses)
    else:
        print("Не удалось извлечь данные.")