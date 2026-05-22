import requests
import json
import re
import time
from openpyxl import Workbook
from urllib.parse import urljoin
from bs4 import BeautifulSoup

BASE_URL = "https://xn--21-1lctt.xn--p1ai"
START_URL = urljoin(BASE_URL, "/specialties.php")
OUTPUT_FILE = "educational_programs.xlsx"
SAVE_EVERY = 20

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}

def get_html(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        r.raise_for_status()
        r.encoding = 'utf-8'
        return r.text
    except Exception as e:
        print(f"  Ошибка загрузки: {e}")
        return None

def extract_specialty_from_html(html):
    """Точно такая же логика, как в проверочном скрипте."""
    # Ищем script type="module"
    script_match = re.search(r'<script\s+type="module">(.*?)</script>', html, re.DOTALL)
    if not script_match:
        return None
    script_content = script_match.group(1)
    
    # Ищем specialty:
    pos = script_content.find('specialty:')
    if pos == -1:
        return None
    
    # Ищем открывающую скобку
    brace_start = script_content.find('{', pos)
    if brace_start == -1:
        return None
    
    # Ручной парсинг с учётом строк и вложенности
    brace_count = 0
    i = brace_start
    in_string = False
    escape = False
    while i < len(script_content):
        ch = script_content[i]
        if not escape:
            if ch == '"' and not in_string:
                in_string = True
            elif ch == '"' and in_string:
                in_string = False
            elif ch == '\\' and in_string:
                escape = True
            elif not in_string:
                if ch == '{':
                    brace_count += 1
                elif ch == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        break
        else:
            escape = False
        i += 1
    
    if brace_count != 0:
        return None
    
    json_str = script_content[brace_start:i+1]
    # Замена Python None на null
    json_str = json_str.replace(': None', ': null')
    try:
        return json.loads(json_str)
    except json.JSONDecodeError as e:
        print(f"  Ошибка JSON: {e}")
        return None

def main():
    print("Загрузка главной страницы...")
    html = get_html(START_URL)
    if not html:
        print("Не удалось загрузить главную страницу")
        return
    
    soup = BeautifulSoup(html, 'lxml')
    accordions = soup.find_all('div', class_='accordion-item')
    print(f"Найдено категорий: {len(accordions)}")
    
    # Сбор всех программ
    all_programs = []
    for accordion in accordions:
        title_elem = accordion.find('div', class_='accordion-title')
        category = title_elem.get_text(strip=True) if title_elem else "Без категории"
        prog_list = accordion.find('div', class_='sub-programs-list')
        if not prog_list:
            continue
        for a in prog_list.find_all('a', class_='sub-program-item', href=True):
            code_elem = a.find('span', class_='sub-program-code')
            name_elem = a.find('span', class_='sub-program-name')
            code = code_elem.get_text(strip=True) if code_elem else ""
            name = name_elem.get_text(strip=True) if name_elem else ""
            url = urljoin(BASE_URL, a['href'])
            all_programs.append((category, code, name, url))
    
    print(f"Всего программ: {len(all_programs)}")
    
    headers = [
        "Категория", "Код программы", "Наименование", "Ссылка",
        "Общая длительность (часы)", "Теория (часы)", "Практика (часы)",
        "Стоимость очного обучения", "Стоимость дистанционного обучения"
    ]
    results = []
    
    for idx, (cat, code, name, url) in enumerate(all_programs, start=1):
        print(f"[{idx}/{len(all_programs)}] {code} - {name}")
        
        # Загружаем страницу программы
        page_html = get_html(url)
        if not page_html:
            results.append([cat, code, name, url, "", "", "", "", ""])
            continue
        
        # Извлекаем JSON
        spec = extract_specialty_from_html(page_html)
        if not spec:
            print(f"  Не удалось извлечь JSON для {url}")
            results.append([cat, code, name, url, "", "", "", "", ""])
            continue
        
        # Извлекаем нужные поля
        total_hours = spec.get('hours', '')
        theory = spec.get('theory_learning', '')
        practice = spec.get('practice_learning', '')
        cost_offline = spec.get('cost', '')
        cost_online = spec.get('cost_distant', '')
        
        # Приводим к строке (чтобы в Excel попадали числа)
        print(f"  -> Часы: {total_hours}, теория: {theory}, практика: {practice}, очно: {cost_offline}, дист: {cost_online}")
        
        results.append([cat, code, name, url, total_hours, theory, practice, cost_offline, cost_online])
        
        # Промежуточное сохранение
        if idx % SAVE_EVERY == 0 or idx == len(all_programs):
            wb = Workbook()
            ws = wb.active
            ws.append(headers)
            for row in results:
                ws.append(row)
            wb.save(OUTPUT_FILE)
            print(f"  >>> Сохранено {idx} записей в {OUTPUT_FILE}")
        
        time.sleep(0.3)
    
    print(f"\nГотово! Обработано {len(results)} программ. Данные сохранены в {OUTPUT_FILE}")

if __name__ == "__main__":
    main()