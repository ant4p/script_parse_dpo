import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import urllib3
import os
import signal
import sys

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://www.ucheba.ru"
SEARCH_URL_BASE = "/for-specialists/search/engineering-and-technology"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'Referer': 'https://www.ucheba.ru/for-specialists/search/engineering-and-technology',
    'Cache-Control': 'max-age=0',
}
DELAY = 3
OUTPUT_FILENAME = 'ucheba_programs.xlsx'

all_data = []  # глобальная переменная для сохранения при прерывании

def signal_handler(sig, frame):
    print("\n⚠️ Прерывание. Сохраняем собранные данные...")
    if all_data:
        save_data(all_data)
    sys.exit(0)

def save_data(data):
    """Сохраняет данные в Excel, объединяя с существующими"""
    if not data:
        return
    if os.path.exists(OUTPUT_FILENAME):
        try:
            existing_df = pd.read_excel(OUTPUT_FILENAME)
            existing_data = existing_df.to_dict('records')
        except Exception as e:
            print(f"Не удалось загрузить существующий файл: {e}")
            existing_data = []
    else:
        existing_data = []

    combined = existing_data + data
    df = pd.DataFrame(combined)
    df = df.drop_duplicates(subset=['Ссылка на программу'], keep='first')
    # Переупорядочиваем колонки: Наименование программы, Ссылка на программу, Учебное заведение, Ссылка на организацию, Стоимость, Длительность
    cols_order = [
        'Наименование программы',
        'Ссылка на программу',
        'Учебное заведение',
        'Ссылка на организацию',
        'Стоимость',
        'Длительность'
    ]
    df = df[cols_order]
    df = df.sort_values('Наименование программы')
    df.to_excel(OUTPUT_FILENAME, index=False)
    print(f"Сохранено {len(df)} записей в {OUTPUT_FILENAME}")

def extract_cost(value_str):
    if not value_str or value_str == "Н/Д":
        return "Н/Д"
    cleaned = re.sub(r'[^\d\s]', '', value_str)
    numbers = re.findall(r'\d+', cleaned)
    return numbers[0] if numbers else "Н/Д"

def parse_page(url):
    print(f"Парсинг: {url}")
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15, verify=False)
        resp.raise_for_status()
    except Exception as e:
        print(f"Ошибка загрузки: {e}")
        return []

    soup = BeautifulSoup(resp.text, 'html.parser')
    cards = soup.find_all('div', class_='cards-list__item')
    if not cards:
        cards = soup.find_all('div', class_='card')
    if not cards:
        print("Карточки не найдены.")
        with open('debug_ucheba.html', 'w', encoding='utf-8') as f:
            f.write(resp.text)
        return []

    page_data = []
    for card in cards:
        # Ссылка и название программы
        link_tag = card.find('a', href=True, class_=lambda x: x and 'js_webstat' in str(x))
        if not link_tag:
            title_tag = card.find('h3', class_='card__title')
            if title_tag:
                link_tag = title_tag.find('a', href=True)
        if not link_tag:
            continue
        program_name = link_tag.get_text(strip=True)
        program_link = link_tag['href']
        if program_link.startswith('/'):
            program_link = BASE_URL + program_link

        # ВУЗ и ссылка на организацию
        vuz_name = ""
        vuz_link = ""
        vuz_block = card.find('dl', class_='card__uz')
        if vuz_block:
            dt = vuz_block.find('dt')
            if dt:
                a = dt.find('a')
                if a:
                    vuz_name = a.get_text(strip=True)
                    vuz_link = a['href']
                    if vuz_link.startswith('/'):
                        vuz_link = BASE_URL + vuz_link

        # Стоимость и длительность
        cost = "Н/Д"
        duration = "Н/Д"
        params_block = card.find('div', class_='card__params')
        if params_block:
            for dl in params_block.find_all('dl'):
                dt = dl.find('dt')
                if not dt:
                    continue
                dt_text = dt.get_text(strip=True).lower()
                dd = dl.find('dd')
                if not dd:
                    continue
                dd_text = dd.get_text(strip=True)
                if 'стоимость' in dt_text:
                    cost = extract_cost(dd_text)
                elif 'длительность' in dt_text:
                    duration = dd_text

        page_data.append({
            'Наименование программы': program_name,
            'Ссылка на программу': program_link,
            'Учебное заведение': vuz_name,
            'Ссылка на организацию': vuz_link,
            'Стоимость': cost,
            'Длительность': duration
        })

    print(f"Найдено программ: {len(page_data)}")
    return page_data

def main():
    global all_data
    signal.signal(signal.SIGINT, signal_handler)

    s = 0          # для ucheba.ru полные карточки начинаются с s=60
    step = 30
    empty_pages = 0

    while True:
        url = f"{BASE_URL}{SEARCH_URL_BASE}?s={s}"
        data = parse_page(url)
        if not data:
            empty_pages += 1
            if empty_pages >= 1:
                print("Нет данных – завершаем.")
                break
        else:
            empty_pages = 0
            all_data.extend(data)
            print(f"Всего собрано: {len(all_data)}")
            if len(data) < 30:
                print("Меньше 30 программ – достигнут конец списка.")
                break
        s += step
        time.sleep(DELAY)

    if all_data:
        save_data(all_data)
    else:
        print("Данные не собраны.")

if __name__ == '__main__':
    main()