import requests
from bs4 import BeautifulSoup
import pandas as pd
import urllib3
import re
from urllib.parse import urljoin
import time

# --- ОТКЛЮЧЕНИЕ SSL-ПРЕДУПРЕЖДЕНИЙ ---
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://energetik-nn.ru"
START_URL = BASE_URL + "/"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
}

def get_soup(url):
    """Загружает страницу и возвращает BeautifulSoup."""
    try:
        response = requests.get(url, headers=HEADERS, verify=False, timeout=10)
        response.encoding = 'utf-8'
        return BeautifulSoup(response.text, 'html.parser')
    except Exception as e:
        print(f"  Ошибка загрузки {url}: {e}")
        return None

def extract_number(text):
    """Извлекает первое число из строки."""
    match = re.search(r'(\d+)', text)
    return match.group(1) if match else ''

def parse_detail_page(url):
    """
    Парсит детальную страницу программы.
    Возвращает словарь с ценой и форматом обучения.
    """
    soup = get_soup(url)
    if not soup:
        return {'price': '', 'format': ''}

    price = ''
    study_format = ''

    # Ищем главный блок с деталями (bannermain1)
    detail_block = soup.find('div', class_='bannermain1')
    if not detail_block:
        # Если блок не найден, пробуем поискать по всей странице
        detail_block = soup

    # Ищем таблицу с классом td-dist (или любую таблицу внутри блока)
    table = detail_block.find('table', class_='td-dist') if detail_block else None
    if not table:
        # Если таблицы нет, ищем просто по тексту
        pass

    # 1. Поиск стоимости
    # Ищем элемент, содержащий текст "Стоимость обучения"
    price_container = None
    if table:
        # Ищем td, внутри которого есть span с текстом "Стоимость обучения"
        for td in table.find_all('td'):
            span = td.find('span', string=re.compile(r'Стоимость обучения', re.I))
            if span:
                # Следующий span (или p) с ценой
                price_span = td.find('span', string=re.compile(r'\d+[\s\d]*руб', re.I))
                if price_span:
                    price = price_span.get_text(strip=True)
                else:
                    # Ищем любой текст с цифрами и руб.
                    price_text = td.get_text(separator=' ', strip=True)
                    match = re.search(r'(\d+[\s\d]*руб\.?)', price_text, re.I)
                    if match:
                        price = match.group(1)
                break

    # Если не нашли через таблицу, ищем по всему тексту
    if not price:
        # Поиск по фразе "Стоимость обучения"
        for tag in soup.find_all(string=re.compile(r'Стоимость обучения', re.I)):
            parent = tag.find_parent()
            if parent:
                # Ищем ближайший span с числом и руб
                price_span = parent.find('span', string=re.compile(r'\d+[\s\d]*руб', re.I))
                if price_span:
                    price = price_span.get_text(strip=True)
                    break

    # Если всё ещё нет, ищем по тексту "руб"
    if not price:
        price_tag = soup.find(string=re.compile(r'\d+[\s\d]*руб', re.I))
        if price_tag:
            price = price_tag.strip()

    # 2. Поиск формата обучения
    # Ищем строку с текстом "Формат обучения:"
    if table:
        for td in table.find_all('td'):
            text = td.get_text(separator=' ', strip=True)
            if 'Формат обучения' in text:
                # Извлекаем текст после ":"
                match = re.search(r'Формат обучения\s*:\s*(.+)', text, re.I)
                if match:
                    study_format = match.group(1).strip()
                else:
                    # Если не нашли через двоеточие, просто берем весь текст ячейки
                    study_format = text
                break

    # Если не нашли в таблице, ищем по всему тексту страницы
    if not study_format:
        for tag in soup.find_all(string=re.compile(r'Формат обучения', re.I)):
            parent = tag.find_parent()
            if parent:
                full_text = parent.get_text(separator=' ', strip=True)
                match = re.search(r'Формат обучения\s*:\s*(.+)', full_text, re.I)
                if match:
                    study_format = match.group(1).strip()
                    break

    # Если всё ещё нет, ищем по ключевым словам "очно", "заочно", "дистанционно"
    if not study_format:
        for tag in soup.find_all(string=re.compile(r'очно|заочно|дистанц', re.I)):
            # Проверяем, что это не просто упоминание, а отдельная фраза
            text = tag.strip()
            if len(text) < 30 and any(word in text.lower() for word in ['очно', 'заочно', 'дистанц']):
                study_format = text
                break

    return {'price': price, 'format': study_format}

def parse_main_page():
    """Парсит главную страницу и собирает все программы."""
    soup = get_soup(START_URL)
    if not soup:
        print("Не удалось загрузить главную страницу.")
        return []

    # Соответствие индекса вкладки и названия вида обучения
    category_names = [
        "Профессиональная подготовка",
        "Профессиональная переподготовка",
        "Повышение квалификации",
        "Дополнительное профессиональное обучение"
    ]

    # Находим все четыре блока вкладок
    tab_blocks = soup.select('.first, .second, .third, .fourth')
    all_programs = []

    for idx, block in enumerate(tab_blocks):
        if idx >= len(category_names):
            break
        category = category_names[idx]
        print(f"Обработка категории: {category}")

        # Внутри блока ищем все карточки программ
        items = block.select('.col-xs-6.col-sm-3')
        for item in items:
            link_tag = item.find('a', href=True)
            if not link_tag:
                continue

            href = link_tag['href']
            full_url = urljoin(BASE_URL, href)

            name_elem = item.select_one('.name_progs')
            name = name_elem.get_text(strip=True) if name_elem else ''

            time_elem = item.select_one('.time_progs')
            hours_text = time_elem.get_text(strip=True) if time_elem else ''
            hours = extract_number(hours_text)

            # Пропускаем ссылки на PDF (не детальные страницы)
            if not href.startswith('/klientam/perechen-kursov.html'):
                print(f"  Пропускаем ссылку (не программа): {href}")
                all_programs.append({
                    'Наименование': name,
                    'Вид обучения': category,
                    'Количество часов': hours,
                    'Стоимость': '',
                    'Формат обучения': '',
                    'Ссылка': full_url
                })
                continue

            print(f"  Обработка: {name}")

            # Парсим детальную страницу
            detail = parse_detail_page(full_url)
            price = detail['price']
            study_format = detail['format']

            all_programs.append({
                'Наименование': name,
                'Вид обучения': category,
                'Количество часов': hours,
                'Стоимость': price,
                'Формат обучения': study_format,
                'Ссылка': full_url
            })

            time.sleep(0.5)  # пауза

    return all_programs

def main():
    print("Начинаем сбор данных с сайта...")
    programs = parse_main_page()

    if programs:
        df = pd.DataFrame(programs)
        df.to_excel('educational_programs.xlsx', index=False)
        print(f"\n✅ Готово! Собрано {len(programs)} программ. Данные сохранены в educational_programs.xlsx")
    else:
        print("❌ Не удалось найти ни одной программы.")

if __name__ == "__main__":
    main()