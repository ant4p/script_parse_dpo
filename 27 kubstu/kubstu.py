import time
import re
import requests
import urllib3
from urllib.parse import urljoin
from bs4 import BeautifulSoup
import pandas as pd

# Отключаем предупреждения SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://kubstu.ru/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
}

def clean_dpo_prefix(text):
    """Убирает из строки префиксы 'Программа ДПО:' или 'Программы ДПО:' (регистронезависимо)."""
    if not text:
        return text
    # удаляем префикс с возможными пробелами после двоеточия
    cleaned = re.sub(r'^Программы?\s+ДПО:\s*', '', text, flags=re.IGNORECASE)
    return cleaned.strip()

def get_soup(url):
    """Загружает страницу, игнорируя ошибки SSL."""
    resp = requests.get(url, headers=HEADERS, verify=False, timeout=10)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or 'utf-8'
    return BeautifulSoup(resp.text, "html.parser")

def parse_program_details(program_url):
    """Парсит детальную страницу программы."""
    soup = get_soup(program_url)
    article = soup.find("article", class_="mside")
    if not article:
        return {}

    # Название программы — очищаем от "Программа ДПО:"
    title_tag = article.find("h1")
    title = title_tag.get_text(strip=True) if title_tag else ""
    title = clean_dpo_prefix(title)

    # Таблица с характеристиками
    table = article.find("table", class_="sb")
    details = {
        "Форма обучения": "",
        "Объем часов": "",
        "Выдаваемый документ": "",
        "Стоимость обучения": "",
    }
    if table:
        for row in table.find_all("tr"):
            th = row.find("th")
            td = row.find("td")
            if not th or not td:
                continue
            key = th.get_text(strip=True)
            value = td.get_text(strip=True)
            if "Форма обучения" in key:
                details["Форма обучения"] = value
            elif "Объем часов" in key:
                details["Объем часов"] = value
            elif "Выдаваемый документ" in key:
                details["Выдаваемый документ"] = value
            elif "Стоимость" in key:
                details["Стоимость обучения"] = value

    return {
        "Наименование программы": title,
        "Ссылка": program_url,
        "Форма обучения": details["Форма обучения"],
        "Объем часов": details["Объем часов"],
        "Выдаваемый документ": details["Выдаваемый документ"],
        "Стоимость обучения": details["Стоимость обучения"],
    }

def parse_all():
    """Главная функция обхода."""
    all_data = []

    # 1. Главная страница с направлениями
    print("Загружаем главную страницу...")
    main_soup = get_soup(BASE_URL + "s-22")
    container = main_soup.find("div", class_="dpo-xpelist-box")
    if not container:
        print("Не найден контейнер с направлениями.")
        return []

    # Собираем все ссылки на направления (group-*)
    direction_links = []
    for a in container.find_all("a", href=True):
        href = a["href"]
        if href.startswith("info/xpelist/group-"):
            full_url = urljoin(BASE_URL, href)
            direction_name = a.get_text(strip=True)
            direction_links.append((full_url, direction_name))

    print(f"Найдено направлений: {len(direction_links)}")

    # 2. Проходим по каждому направлению
    for dir_url, dir_name in direction_links:
        print(f"\nОбрабатываем: {dir_name}")
        try:
            soup = get_soup(dir_url)
        except Exception as e:
            print(f"  Ошибка загрузки {dir_url}: {e}")
            continue

        article = soup.find("article", class_="mside")
        if not article:
            continue

        # Название раздела (факультет) — очищаем от "Программы ДПО:"
        header = article.find("h1")
        section_name = header.get_text(strip=True) if header else dir_name
        section_name = clean_dpo_prefix(section_name)   # <-- очистка

        # Ищем вложенные списки программ
        main_ul = article.find("ul", class_="lxm")
        if not main_ul:
            continue

        for li in main_ul.find_all("li", recursive=False):
            dept_tag = li.find("b")
            if not dept_tag:
                continue
            dept_name = dept_tag.get_text(strip=True)

            inner_ul = li.find("ul")
            if not inner_ul:
                continue

            for program_li in inner_ul.find_all("li"):
                a_tag = program_li.find("a", href=True)
                if not a_tag:
                    continue
                program_url = urljoin(BASE_URL, a_tag["href"])
                program_title = a_tag.get_text(strip=True)

                # Получаем детали программы
                try:
                    details = parse_program_details(program_url)
                except Exception as e:
                    print(f"    Ошибка парсинга {program_url}: {e}")
                    continue

                if not details:
                    continue

                # Добавляем раздел (факультет / кафедра) — здесь не нужно чистить, потому что section_name уже чистый
                details["Раздел"] = f"{section_name} / {dept_name}"
                if not details["Наименование программы"]:
                    details["Наименование программы"] = program_title

                all_data.append(details)
                print(f"    Добавлена: {details['Наименование программы'][:50]}...")

                # Небольшая задержка, чтобы не перегружать сервер
                time.sleep(0.3)

    return all_data

def save_to_excel(data, filename="programs.xlsx"):
    if not data:
        print("Нет данных для сохранения.")
        return

    df = pd.DataFrame(data)
    columns_order = [
        "Наименование программы",
        "Ссылка",
        "Раздел",
        "Форма обучения",
        "Объем часов",
        "Выдаваемый документ",
        "Стоимость обучения",
    ]
    df = df[columns_order]
    df.to_excel(filename, index=False)
    print(f"\n✅ Данные сохранены в {filename}")

if __name__ == "__main__":
    data = parse_all()
    save_to_excel(data)