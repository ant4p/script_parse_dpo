import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin
import time
import urllib3

# Отключаем предупреждения о небезопасном SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = 'https://npi-tu.ru'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def get_soup(url):
    """Загружает страницу и возвращает BeautifulSoup, игнорируя SSL-ошибки."""
    try:
        response = requests.get(url, headers=HEADERS, timeout=15, verify=False)
        response.raise_for_status()
        # Если текст отображается некорректно, попробуйте раскомментировать:
        # response.encoding = 'windows-1251'
        return BeautifulSoup(response.text, 'html.parser')
    except requests.exceptions.RequestException as e:
        print(f'Ошибка при загрузке {url}: {e}')
        return None

def get_category_links(main_url):
    """Извлекает все ссылки на категории с главной страницы."""
    soup = get_soup(main_url)
    if not soup:
        return []
    
    main_block = soup.find('main', class_='sf-main-area')
    if not main_block:
        print('Не найден блок main.sf-main-area')
        return []
    
    ul = main_block.find('ul')
    if not ul:
        print('Не найден <ul> внутри main')
        return []
    
    links = []
    for li in ul.find_all('li'):
        a = li.find('a')
        if a and a.get('href'):
            href = a['href']
            full_url = urljoin(BASE_URL, href)
            category_name = a.get_text(strip=True)
            links.append((category_name, full_url))
    
    return links

def parse_program_page(category_name, page_url):
    """Парсит страницу категории. Адаптируйте под реальную структуру."""
    soup = get_soup(page_url)
    if not soup:
        return []
    
    programs = []
    
    # Вариант 1: Таблица
    table = soup.find('table')
    if table:
        rows = table.find_all('tr')
        for row in rows[1:]:  # если первая строка — заголовок
            cols = row.find_all('td')
            if len(cols) >= 3:
                name = cols[0].get_text(strip=True)
                hours = cols[1].get_text(strip=True)
                price = cols[2].get_text(strip=True)
                programs.append({
                    'category': category_name,
                    'link': page_url,
                    'name': name,
                    'hours': hours,
                    'price': price
                })
        if programs:
            return programs
    
    # Вариант 2: Блоки с классами (пример)
    items = soup.find_all('div', class_='program-item')
    if items:
        for item in items:
            name_elem = item.find(class_='program-name')
            hours_elem = item.find(class_='program-hours')
            price_elem = item.find(class_='program-price')
            if name_elem:
                programs.append({
                    'category': category_name,
                    'link': page_url,
                    'name': name_elem.get_text(strip=True),
                    'hours': hours_elem.get_text(strip=True) if hours_elem else '',
                    'price': price_elem.get_text(strip=True) if price_elem else ''
                })
        if programs:
            return programs
    
    # Вариант 3: список <li> с текстом (если данные не в таблице и не в блоках)
    # Например, каждая программа в <li>, а внутри текст вида "Название – 72 ч – 15000 руб"
    # Тогда можно распарсить регуляркой или по разделителям.
    # Здесь нужно добавить свою логику, если структура иная.
    
    print(f'Не удалось распарсить {page_url} — структура не соответствует ожидаемой.')
    return []

def main():
    main_url = 'https://npi-tu.ru/university/faculty/ino/tsdo/povyshenie-kvalifikatsii/ds/'
    
    categories = get_category_links(main_url)
    print(f'Найдено категорий: {len(categories)}')
    
    if not categories:
        print('Проверьте доступность сайта или структуру страницы.')
        return
    
    all_data = []
    
    for idx, (cat_name, cat_url) in enumerate(categories, start=1):
        print(f'[{idx}/{len(categories)}] Парсинг: {cat_name} ({cat_url})')
        programs = parse_program_page(cat_name, cat_url)
        if programs:
            all_data.extend(programs)
            print(f'  → Найдено {len(programs)} программ')
        else:
            print(f'  → Программ не найдено')
        time.sleep(1)  # вежливость
    
    if all_data:
        df = pd.DataFrame(all_data)
        df = df[['category', 'link', 'name', 'hours', 'price']]
        excel_file = 'programs.xlsx'
        df.to_excel(excel_file, index=False, engine='openpyxl')
        print(f'✅ Сохранено {len(df)} записей в {excel_file}')
    else:
        print('❌ Нет данных для сохранения.')

if __name__ == '__main__':
    main()