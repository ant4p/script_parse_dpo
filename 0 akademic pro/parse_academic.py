import requests
from bs4 import BeautifulSoup
import pandas as pd
import urllib3

# Отключаем предупреждения о сертификатах
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URLS = [
    ("https://akademik-pro.ru/nashi_kursy/kursy_povysheniya_kvalifikacii/", "Курс"),
    ("https://akademik-pro.ru/nashi_kursy/professionalnyie_standarty.html", "Стандарт")
]
HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}

def parse_table(soup, category):
    """
    Парсит таблицу с классом 'main_table' и добавляет категорию.
    Возвращает список строк: [номер, название, должности, часы, категория]
    """
    table = soup.find('table', class_='main_table')
    if not table:
        return []

    rows = table.find_all('tr')
    if len(rows) < 2:
        return []

    data = []
    for row in rows[1:]:
        cells = row.find_all(['td', 'th'])
        cell_texts = [cell.get_text(strip=True) for cell in cells if cell.get_text(strip=True)]
        if not cell_texts:
            continue

        num_cols = len(cell_texts)
        if num_cols == 3:
            # Страница курсов: №, Название, Часы
            num, name, hours = cell_texts[0], cell_texts[1], cell_texts[2]
            duties = ''
        elif num_cols == 4:
            # Страница стандартов: №, Название, Должности, Часы
            num, name, duties, hours = cell_texts[0], cell_texts[1], cell_texts[2], cell_texts[3]
        else:
            continue

        # Очистка номера и часов
        num = num.rstrip('.')
        hours = hours.replace('—', '').strip()
        data.append([num, name, duties, hours, category])

    return data

def main():
    all_data = []

    for url, category in URLS:
        print(f"🚀 Загрузка {url} (категория: {category})")
        try:
            response = requests.get(url, headers=HEADERS, timeout=15, verify=False)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            page_data = parse_table(soup, category)
            if page_data:
                all_data.extend(page_data)
                print(f"   ✅ Найдено {len(page_data)} записей")
            else:
                print("   ⚠️ Данных не найдено")

        except Exception as e:
            print(f"   ❌ Ошибка: {e}")

    if not all_data:
        print("❌ Нет данных для сохранения")
        return

    # Создаём DataFrame с новым столбцом "Категория"
    df = pd.DataFrame(all_data, columns=['№ п/п', 'Наименование программы', 'Должности', 'Кол-во часов', 'Категория'])
    filename = 'akademik_pro_courses.xlsx'
    df.to_excel(filename, index=False)
    print(f"✅ Сохранено {len(df)} записей в {filename}")

if __name__ == "__main__":
    main()