import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin
import re
from urllib3.exceptions import InsecureRequestWarning

# Отключаем предупреждения о небезопасном SSL-соединении
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

# --- Настройки ---
BASE_URL = "https://ecoips.ru"
MAIN_URL = "https://ecoips.ru/obuchenie-i-uslugi/"

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

# --- Функции для извлечения данных ---
def get_soup(url):
    """Загружает страницу и возвращает объект BeautifulSoup."""
    try:
        # verify=False - отключаем проверку SSL-сертификата
        response = requests.get(url, headers=headers, timeout=10, verify=False)
        response.raise_for_status()
        response.encoding = 'utf-8'
        return BeautifulSoup(response.text, 'html.parser')
    except requests.exceptions.RequestException as e:
        print(f"Ошибка загрузки {url}: {e}")
        return None

def get_program_hours(program_url):
    """Извлекает количество часов со страницы программы."""
    soup = get_soup(program_url)
    if not soup:
        return None

    # Ищем текст "Продолжительность курса:" и извлекаем число
    for text_elem in soup.find_all(text=True):
        if "Продолжительность курса:" in text_elem:
            parent = text_elem.find_parent()
            if parent:
                text = parent.get_text()
                match = re.search(r'Продолжительность курса:\s*(\d+)', text)
                if match:
                    return match.group(1)
            break
    return None

# --- Основной цикл сбора данных ---
all_courses = []

print("Шаг 1: Парсинг главной страницы...")
main_soup = get_soup(MAIN_URL)
if not main_soup:
    print("Не удалось загрузить главную страницу.")
    exit()

# 1. Находим все блоки направлений (разделы)
sections = main_soup.select('.box a.item')
print(f"Найдено разделов: {len(sections)}")

for section in sections:
    section_name_elem = section.select_one('.name')
    if not section_name_elem:
        continue
    section_name = section_name_elem.get_text(strip=True)
    section_url = urljoin(BASE_URL, section.get('href'))
    print(f"\nОбработка раздела: {section_name} ({section_url})")

    # 2. Загружаем страницу раздела
    section_soup = get_soup(section_url)
    if not section_soup:
        continue

    # 3. Находим все ссылки на программы внутри #content
    programs_container = section_soup.find('div', id='content')
    if not programs_container:
        print(f"Не найден контейнер #content в разделе {section_name}")
        continue

    program_links = programs_container.select('a.item')
    print(f"Найдено программ в разделе: {len(program_links)}")

    for program in program_links:
        program_name_elem = program.select_one('.zag')
        hours_elem = program.select_one('.date')

        if not program_name_elem:
            continue

        program_name = program_name_elem.get_text(strip=True)
        hours_on_page = hours_elem.get_text(strip=True) if hours_elem else None
        program_url = urljoin(BASE_URL, program.get('href'))

        # 4. Если часы не указаны на странице списка, идём на страницу программы
        program_hours = hours_on_page
        if not program_hours and program_url:
            program_hours = get_program_hours(program_url)

        all_courses.append({
            'Раздел': section_name,
            'Программа': program_name,
            'Часы': program_hours,
            'Ссылка': program_url
        })

# --- Сохранение результатов ---
print(f"\n\nСобрано записей: {len(all_courses)}")
df = pd.DataFrame(all_courses)
df.to_excel('educational_programs.xlsx', index=False)
print("Данные сохранены в файл 'educational_programs.xlsx'")