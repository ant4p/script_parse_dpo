import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin
import time
import urllib3
import re

# Отключаем предупреждения SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = 'https://ipo.msk.ru'
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

# Разделы для парсинга
SECTIONS = [
    {'name': 'Профессиональная переподготовка', 'url': '/professionalnaja-perepodgotovka/'},
    {'name': 'Повышение квалификации', 'url': '/povyshenie-kvalifikacii/'}
]

def get_soup(url):
    """Загружает страницу и возвращает BeautifulSoup."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15, verify=False)
        resp.raise_for_status()
        # Если сайт использует другую кодировку, раскомментируйте:
        # resp.encoding = 'windows-1251'
        return BeautifulSoup(resp.text, 'html.parser')
    except Exception as e:
        print(f'Ошибка загрузки {url}: {e}')
        return None

def get_category_links(section_url):
    """Собирает категории из раздела (блок .list-prog-dpo ul.nily-field-items)."""
    full_url = urljoin(BASE_URL, section_url)
    soup = get_soup(full_url)
    if not soup:
        return []
    
    ul = soup.select_one('div.list-prog-dpo ul.nily-field-items')
    if not ul:
        print(f'Не найден блок категорий на {full_url}')
        return []
    
    categories = []
    for li in ul.find_all('li', class_='nily-field-items__nily-item'):
        a = li.find('a', class_='nily-item__nily-link')
        if a and a.get('href'):
            href = a['href']
            full_cat_url = urljoin(BASE_URL, href)
            name_elem = a.find('h3', class_='nily-link__program-name')
            cat_name = name_elem.get_text(strip=True) if name_elem else 'Без названия'
            categories.append({'name': cat_name, 'url': full_cat_url})
    return categories

def get_program_links(category_url):
    """
    Извлекает ссылки на программы со страницы категории.
    Поддерживает два формата:
    1) <div class="list-prog-dpo"> -> <ul> -> <li><a>
    2) <div class="program-list"> -> <a class="program-list__program-title">
    """
    soup = get_soup(category_url)
    if not soup:
        return []
    
    program_urls = []
    
    # ---- Тип 1: блок .list-prog-dpo (агрономия, нефтегаз и др.) ----
    list_dpo = soup.find('div', class_='list-prog-dpo')
    if list_dpo:
        ul = list_dpo.find('ul')
        if ul:
            for li in ul.find_all('li'):
                a = li.find('a', href=True)
                if a:
                    full_url = urljoin(BASE_URL, a['href'])
                    program_urls.append(full_url)
            if program_urls:
                return list(set(program_urls))
    
    # ---- Тип 2: блок .program-list (экономика и др.) ----
    program_list = soup.find('div', class_='program-list')
    if program_list:
        for a in program_list.find_all('a', class_='program-list__program-title', href=True):
            full_url = urljoin(BASE_URL, a['href'])
            program_urls.append(full_url)
        if program_urls:
            return list(set(program_urls))
    
    return program_urls

def parse_program_page(program_url):
    """
    Парсит страницу одной программы.
    Извлекает название, часы, длительность, цены.
    """
    soup = get_soup(program_url)
    if not soup:
        return {}
    
    data = {'url': program_url}
    
    # 1. Название (обычно в <h1>)
    h1 = soup.find('h1')
    data['name'] = h1.get_text(strip=True) if h1 else ''
    
    # 2. Вид и формат (если есть – можно донастроить)
    data['type'] = ''
    data['format'] = ''
    
    # 3. Блок с ценами и часами
    promo = soup.find('div', class_='promo-block')
    if promo:
        # ---- Часы и длительность ----
        # Ищем все элементы с классом promo-block-item-list
        list_blocks = promo.find_all('div', class_='promo-block-item-list')
        hours = ''
        duration = ''
        for block in list_blocks:
            text = block.get_text(separator=' ', strip=True)  # склеиваем через пробел
            # Ищем часы: число + "час" или "ч."
            match = re.search(r'(\d+)\s*(?:час|ч\.)', text, re.IGNORECASE)
            if match:
                hours = match.group(1) + ' ч.'
            # Ищем длительность: "до X месяца" или "до X месяцев"
            dur_match = re.search(r'до\s*(\d+)\s*месяц', text, re.IGNORECASE)
            if dur_match:
                duration = f"до {dur_match.group(1)} месяца"
            # Если нашли и часы, и длительность – можно выйти, но продолжаем на случай, если в разных блоках
        data['hours'] = hours
        data['duration'] = duration
        
        # ---- Цены ----
        price_new = promo.find('div', class_='promo-block-item-pnew')
        data['price_new'] = price_new.get_text(strip=True) if price_new else ''
        
        price_old = promo.find('div', class_='promo-block-item-pold')
        data['price_old'] = price_old.get_text(strip=True) if price_old else ''
    else:
        data['hours'] = ''
        data['duration'] = ''
        data['price_new'] = ''
        data['price_old'] = ''
    
    return data

def main():
    all_data = []
    
    for section in SECTIONS:
        print(f'Обработка раздела: {section["name"]}')
        categories = get_category_links(section['url'])
        print(f'  Найдено категорий: {len(categories)}')
        
        for cat in categories:
            print(f'    Категория: {cat["name"]} ({cat["url"]})')
            program_links = get_program_links(cat['url'])
            print(f'      Найдено программ: {len(program_links)}')
            
            for p_url in program_links:
                print(f'        Парсинг: {p_url}')
                prog_data = parse_program_page(p_url)
                if prog_data and prog_data.get('name'):  # проверяем, что есть название
                    prog_data['section'] = section['name']
                    prog_data['category'] = cat['name']
                    all_data.append(prog_data)
                else:
                    print(f'          ⚠️ Не удалось получить данные')
                time.sleep(0.5)  # вежливая пауза
    
    if all_data:
        df = pd.DataFrame(all_data)
        # Упорядочиваем колонки
        columns = ['section', 'category', 'name', 'type', 'format', 'hours', 'duration', 'price_new', 'price_old', 'url']
        df = df[columns]
        df.to_excel('programs_ipo.xlsx', index=False, engine='openpyxl')
        print(f'✅ Сохранено {len(df)} записей в programs_ipo.xlsx')
    else:
        print('❌ Нет данных. Проверьте селекторы.')

if __name__ == '__main__':
    main()