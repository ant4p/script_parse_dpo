import requests
from bs4 import BeautifulSoup
import pandas as pd
import urllib3
import re
from urllib.parse import urljoin
import time

# --- ОТКЛЮЧЕНИЕ ПРЕДУПРЕЖДЕНИЙ SSL ---
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://lichtnostniyrost.ru"
START_URL = urljoin(BASE_URL, "/online-education/")

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def get_soup(url, retries=3):
    """Загружает страницу, возвращает BeautifulSoup или None."""
    for attempt in range(retries):
        try:
            response = requests.get(url, headers=HEADERS, verify=False, timeout=10)
            response.encoding = 'windows-1251'
            return BeautifulSoup(response.text, 'html.parser')
        except Exception as e:
            print(f"Попытка {attempt+1}/{retries} загрузить {url} не удалась: {e}")
            time.sleep(2)
    return None

def parse_program_detail(program_url, default_name, default_hours):
    """
    Парсит детальную страницу программы.
    Возвращает словарь с полями: name, for_whom, study_form, hours, price.
    Если поле не найдено, берёт значения по умолчанию.
    """
    soup = get_soup(program_url)
    if not soup:
        return {
            'name': default_name,
            'for_whom': '',
            'study_form': '',
            'hours': default_hours,
            'price': ''
        }

    # Ищем все блоки свойств
    properties = soup.select('.catalog-detail__property')
    for prop in properties:
        title_elem = prop.select_one('.catalog-detail__property-title')
        value_elem = prop.select_one('.catalog-detail__property-value')
        if title_elem and value_elem:
            title = title_elem.get_text(strip=True)
            value = value_elem.get_text(strip=True)
            if 'Для кого' in title:
                for_whom = value
            elif 'Форма обучения' in title:
                study_form = value
            elif 'Объем ак. часов' in title:
                hours = value

    # Стоимость
    price_elem = soup.select_one('.catalog-detail__price')
    price = price_elem.get_text(strip=True) if price_elem else ''

    # Очистка стоимости (удаляем "руб." и пробелы)
    price = re.sub(r'руб\.?', '', price).strip()

    # Если название не найдено на детальной странице, оставляем из списка
    # Пытаемся найти заголовок h1
    title_elem = soup.find('h1')
    name = title_elem.get_text(strip=True) if title_elem else default_name

    return {
        'name': name,
        'for_whom': for_whom if 'for_whom' in locals() else '',
        'study_form': study_form if 'study_form' in locals() else '',
        'hours': hours if 'hours' in locals() else default_hours,
        'price': price
    }

def parse_programs_from_list(soup, category_name, subcategory_name=''):
    """Извлекает все программы из блока catalog-list и собирает детали."""
    programs = []
    items = soup.select('.catalog-list__item')
    for item in items:
        link_elem = item.select_one('.catalog-list__link')
        if not link_elem:
            continue
        href = link_elem.get('href')
        if not href:
            continue
        program_url = urljoin(BASE_URL, href)

        title_elem = item.select_one('.catalog-list__title')
        title_text = title_elem.get_text(strip=True) if title_elem else ''

        # Извлекаем название и часы из заголовка (как запасной вариант)
        # Пример: «Инструктор-реаниматор», 40 ак.часов
        name_match = re.search(r'«(.*?)»', title_text)
        name = name_match.group(1) if name_match else title_text

        hours_match = re.search(r'(\d+)\s*ак\.?часов?', title_text)
        hours = hours_match.group(1) if hours_match else ''

        # Если не нашли часы в заголовке, берем из .catalog-list__length
        if not hours:
            length_elem = item.select_one('.catalog-list__length')
            if length_elem:
                hours = length_elem.get_text(strip=True)

        # Парсим детальную страницу
        detail = parse_program_detail(program_url, name, hours)

        programs.append({
            'Наименование': detail['name'],
            'Категория': category_name,
            'Подкатегория': subcategory_name,
            'Ссылка на программу': program_url,
            'Для кого': detail['for_whom'],
            'Форма обучения': detail['study_form'],
            'Объем ак. часов': detail['hours'],
            'Стоимость': detail['price']
        })
        # Пауза, чтобы не перегружать сервер
        time.sleep(0.5)
    return programs

def parse_category(category_url, category_name):
    """Парсит категорию: ищет подкатегории или сразу программы."""
    soup = get_soup(category_url)
    if not soup:
        return []

    all_programs = []

    # Проверяем наличие подкатегорий
    subcategory_links = soup.select('a.section-list__link_depth_more')
    if subcategory_links:
        # Есть подкатегории – обходим каждую
        for link in subcategory_links:
            sub_name = link.get_text(strip=True)
            sub_url = urljoin(BASE_URL, link.get('href'))
            print(f"  Обработка подкатегории: {sub_name}")
            sub_soup = get_soup(sub_url)
            if sub_soup:
                programs = parse_programs_from_list(sub_soup, category_name, sub_name)
                all_programs.extend(programs)
                print(f"    Найдено программ: {len(programs)}")
    else:
        # Подкатегорий нет – собираем программы прямо на странице категории
        print(f"  Сбор программ в категории (без подкатегорий)")
        programs = parse_programs_from_list(soup, category_name, '')
        all_programs.extend(programs)
        print(f"    Найдено программ: {len(programs)}")

    return all_programs

def main():
    print("Начинаем сбор данных...")
    main_soup = get_soup(START_URL)
    if not main_soup:
        print("Не удалось загрузить главную страницу.")
        return

    # Все категории (первый уровень)
    category_links = main_soup.select('a.section-list__link_depth_1')
    all_data = []

    for link in category_links:
        cat_name = link.get_text(strip=True)
        cat_url = urljoin(BASE_URL, link.get('href'))
        print(f"Обработка категории: {cat_name}")
        programs = parse_category(cat_url, cat_name)
        all_data.extend(programs)
        print(f"  Всего в категории: {len(programs)} программ\n")

    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel('educational_programs.xlsx', index=False)
        print(f"\n✅ Готово! Собрано {len(all_data)} программ. Данные сохранены в educational_programs.xlsx")
    else:
        print("❌ Не удалось найти ни одной программы.")

if __name__ == "__main__":
    main()