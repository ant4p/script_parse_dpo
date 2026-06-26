import requests
from bs4 import BeautifulSoup
import re
import time
from urllib.parse import urljoin
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = 'https://pdo-osnova.ru/'
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}
session = requests.Session()
session.headers.update(HEADERS)

def get_soup(url):
    try:
        response = session.get(url, timeout=15, verify=False)
        response.raise_for_status()
        response.encoding = 'utf-8'
        return BeautifulSoup(response.text, 'html.parser')
    except Exception as e:
        print(f'Ошибка при загрузке {url}: {e}')
        return None

def extract_price(price_text):
    if not price_text:
        return None
    cleaned = re.sub(r'[^\d\s.,]', '', price_text).strip()
    cleaned = re.sub(r'^от\s*', '', cleaned)
    cleaned = cleaned.replace(' ', '').replace(',', '.')
    try:
        return float(cleaned)
    except ValueError:
        return None

def get_sections():
    soup = get_soup(BASE_URL)
    if not soup:
        print('Главная страница не загружена, использую заранее известные разделы')
        known = [
            {'title': 'Повышение квалификации', 'url': urljoin(BASE_URL, 'povyshenie-kvalifikatsii/')},
            {'title': 'Профессиональная переподготовка', 'url': urljoin(BASE_URL, 'professionalnaya-perepodgotovka/')},
            {'title': 'Обучение по рабочим специальностям', 'url': urljoin(BASE_URL, 'obuchenie-rabochim-specialnostyam/')},
            {'title': 'Очное обучение', 'url': urljoin(BASE_URL, 'ochnoe-obuchenie/')},
            {'title': 'Семинары', 'url': urljoin(BASE_URL, 'seminars/')},
            {'title': 'Аттестация персонала', 'url': urljoin(BASE_URL, 'attestaciya/')},
        ]
        return known

    sections = []
    container = soup.find('div', class_='grid-list grid-list--items-2-991')
    if not container:
        containers = soup.find_all('div', class_=re.compile(r'grid-list--items-2'))
        for cont in containers:
            items = cont.find_all('div', class_='banners-img-with-text-list__wrapper')
            if items:
                container = cont
                break

    if container:
        items = container.find_all('div', class_='banners-img-with-text-list__wrapper')
        for item in items:
            link_tag = item.find('a', class_='dark_link')
            if link_tag:
                title = link_tag.get_text(strip=True)
                href = link_tag.get('href')
                if href:
                    full_url = urljoin(BASE_URL, href)
                    sections.append({'title': title, 'url': full_url})
    else:
        print('Контейнер не найден, пробую поискать ссылки вручную')
        known_paths = ['povyshenie-kvalifikatsii', 'professionalnaya-perepodgotovka', 
                       'obuchenie-rabochim-specialnostyam', 'ochnoe-obuchenie', 
                       'seminars', 'attestaciya']
        for path in known_paths:
            link = soup.find('a', href=re.compile(path))
            if link:
                title = link.get_text(strip=True)
                if title:
                    sections.append({'title': title, 'url': urljoin(BASE_URL, link.get('href'))})
        if not sections:
            print('Не удалось найти разделы, использую список по умолчанию')
            sections = [
                {'title': 'Повышение квалификации', 'url': urljoin(BASE_URL, 'povyshenie-kvalifikatsii/')},
                {'title': 'Профессиональная переподготовка', 'url': urljoin(BASE_URL, 'professionalnaya-perepodgotovka/')},
                {'title': 'Обучение по рабочим специальностям', 'url': urljoin(BASE_URL, 'obuchenie-rabochim-specialnostyam/')},
                {'title': 'Очное обучение', 'url': urljoin(BASE_URL, 'ochnoe-obuchenie/')},
                {'title': 'Семинары', 'url': urljoin(BASE_URL, 'seminars/')},
                {'title': 'Аттестация персонала', 'url': urljoin(BASE_URL, 'attestaciya/')},
            ]
    return sections

def get_programs_from_page(soup, section_title):
    programs = []
    items = soup.find_all('div', class_='services-list__wrapper')
    for item in items:
        title_tag = item.find('a', class_='dark_link')
        if not title_tag:
            continue
        title = title_tag.get_text(strip=True)
        href = title_tag.get('href')
        if not href:
            continue
        program_url = urljoin(BASE_URL, href)
        price_tag = item.find('span', class_='price__new-val')
        price_text = price_tag.get_text(strip=True) if price_tag else ''
        price = extract_price(price_text)
        programs.append({
            'section': section_title,
            'name': title,
            'url': program_url,
            'price': price,
            'hours': None,
            'format': None
        })
    return programs

def get_program_details(program_url):
    """Возвращает словарь с часами и форматом обучения для программы."""
    soup = get_soup(program_url)
    if not soup:
        return {'hours': None, 'format': None}

    hours = None
    format_val = None

    # 1. Ищем свойства в props_block и по всему документу
    prop_items = soup.find_all('div', itemprop='additionalProperty')
    for prop in prop_items:
        name_span = prop.find('span', itemprop='name')
        if not name_span:
            continue
        name = name_span.get_text(strip=True)
        value_div = prop.find('div', itemprop='value')
        if not value_div:
            continue
        value = value_div.get_text(strip=True)

        # Часы: сначала Количество часов, затем Сроки обучения
        if hours is None and 'Количество часов' in name:
            hours = value
        if hours is None and 'Сроки обучения' in name:
            hours = value

        # Формат
        if format_val is None and ('Формат обучения' in name or 'Форма обучения' in name or 'Формат' in name):
            format_val = value

    # 2. Если часы не найдены, ищем в таблице описания (блок desc)
    if hours is None:
        desc_block = soup.find('div', class_='detail-block ordered-block desc')
        if desc_block:
            table = desc_block.find('table')
            if table:
                ths = table.find_all('th')
                headers = [th.get_text(strip=True) for th in ths]
                col_idx = None
                for i, h in enumerate(headers):
                    if h.lower() in ('длительность', 'объем', 'часы', 'количество часов', 'срок'):
                        col_idx = i
                        break
                if col_idx is not None:
                    tbody = table.find('tbody')
                    rows = tbody.find_all('tr') if tbody else table.find_all('tr')[1:]
                    for row in rows:
                        cells = row.find_all('td')
                        if len(cells) > col_idx:
                            cell_text = cells[col_idx].get_text(strip=True)
                            if cell_text and ('часов' in cell_text or re.search(r'\d+', cell_text)):
                                hours = cell_text
                                break

    # 3. Если всё ещё нет часов, ищем по тексту "часов" во всём документе
    if hours is None:
        match = re.search(r'(\d+\s*часов)', soup.get_text())
        if match:
            hours = match.group(1)

    # 4. Если формат не найден, ищем в таблице
    if format_val is None:
        desc_block = soup.find('div', class_='detail-block ordered-block desc')
        if desc_block:
            table = desc_block.find('table')
            if table:
                ths = table.find_all('th')
                headers = [th.get_text(strip=True) for th in ths]
                col_idx = None
                for i, h in enumerate(headers):
                    if h.lower() in ('формат', 'форма обучения', 'обучение'):
                        col_idx = i
                        break
                if col_idx is not None:
                    tbody = table.find('tbody')
                    rows = tbody.find_all('tr') if tbody else table.find_all('tr')[1:]
                    for row in rows:
                        cells = row.find_all('td')
                        if len(cells) > col_idx:
                            cell_text = cells[col_idx].get_text(strip=True)
                            if cell_text:
                                format_val = cell_text
                                break

    # 5. Если формат всё ещё не найден, ищем ключевые слова в тексте
    if format_val is None:
        text = soup.get_text()
        format_keywords = ['очное', 'дистанционное', 'заочное', 'очно-заочное', 'вечернее', 'смешанное']
        found = []
        for kw in format_keywords:
            if re.search(r'\b' + kw + r'\b', text, re.IGNORECASE):
                found.append(kw.capitalize())
        if found:
            format_val = ', '.join(found)

    return {'hours': hours, 'format': format_val}

def get_all_programs_in_section(section):
    all_programs = []
    url = section['url']
    page_num = 1
    while True:
        if '?' in url:
            paginated_url = url + f'&PAGEN_155={page_num}'
        else:
            paginated_url = url + f'?PAGEN_155={page_num}'
        print(f'Загрузка страницы {page_num} раздела {section["title"]}: {paginated_url}')
        soup = get_soup(paginated_url)
        if not soup:
            break
        programs = get_programs_from_page(soup, section['title'])
        if not programs:
            break
        all_programs.extend(programs)
        pagination = soup.find('div', class_='module-pagination')
        if pagination:
            next_link = pagination.find('a', class_='arrows-pagination__next')
            if not next_link:
                next_link = pagination.find('a', string=re.compile(r'След\.|>'))
            if next_link:
                page_num += 1
                continue
            else:
                page_links = pagination.find_all('a', class_='module-pagination__item')
                page_numbers = [int(a.get_text(strip=True)) for a in page_links if a.get_text(strip=True).isdigit()]
                if page_numbers and max(page_numbers) >= page_num:
                    page_num += 1
                    continue
        break
        time.sleep(1)
    return all_programs

def main():
    print('Получение списка разделов...')
    sections = get_sections()
    if not sections:
        print('Разделы не найдены.')
        return
    print(f'Найдено разделов: {len(sections)}')
    all_data = []
    for idx, section in enumerate(sections, 1):
        print(f'\nОбработка раздела {idx}/{len(sections)}: {section["title"]}')
        programs = get_all_programs_in_section(section)
        print(f'Найдено программ: {len(programs)}')
        for prog in programs:
            print(f'  Загрузка страницы программы: {prog["name"]}')
            details = get_program_details(prog['url'])
            prog['hours'] = details['hours']
            prog['format'] = details['format']
            time.sleep(0.5)
        all_data.extend(programs)

    df = pd.DataFrame(all_data, columns=['section', 'name', 'url', 'format', 'hours', 'price'])
    if not df.empty:
        df['price'] = df['price'].astype(float).round(2)

    df = df[['section', 'name', 'url', 'format', 'hours', 'price']]

    output_file = 'educational_programs.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f'\nГотово! Данные сохранены в {output_file}')
    print(f'Всего собрано записей: {len(df)}')

if __name__ == '__main__':
    main()