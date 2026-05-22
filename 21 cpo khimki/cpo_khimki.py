import requests
import pandas as pd
import re
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def parse_qualification(soup):
    """Парсит страницу повышения квалификации"""
    data = []
    for element in soup.select("#content .span12"):
        for h2 in element.find_all("h2"):
            category = h2.get_text(strip=True).replace("## ", "")
            next_table = h2.find_next_sibling("div", class_="pricing-tables")
            if next_table:
                rows = next_table.find_all("tr")
                for row in rows[1:]:
                    cols = row.find_all("td")
                    if len(cols) >= 4:
                        data.append({
                            "Категория": category,
                            "Номер п/п": cols[0].get_text(strip=True),
                            "Наименование программы": cols[1].get_text(strip=True),
                            "Количество часов": cols[2].get_text(strip=True),
                            "Код программы": cols[3].get_text(strip=True)
                        })
    return data

def parse_retraining(soup):
    """Парсит страницу профессиональной переподготовки"""
    data = []
    for group in soup.select(".accordion-group"):
        toggle = group.select_one(".accordion-toggle")
        if not toggle:
            continue
        raw_text = toggle.get_text(strip=True)
        # Извлекаем название и часы
        if " — " in raw_text:
            parts = raw_text.rsplit(" — ", 1)
            name = parts[0].strip()
            hours_raw = parts[1].strip()
        elif " - " in raw_text:
            parts = raw_text.rsplit(" - ", 1)
            name = parts[0].strip()
            hours_raw = parts[1].strip()
        else:
            name = raw_text
            hours_raw = ""
        # Чистим часы
        if hours_raw:
            hours_match = re.search(r'^[\d/]+', hours_raw)
            hours = hours_match.group(0) if hours_match else hours_raw
        else:
            hours = ""
        data.append({
            "Категория": "Профессиональная переподготовка",
            "Номер п/п": "",
            "Наименование программы": name,
            "Количество часов": hours,
            "Код программы": ""
        })
    return data

def fetch_soup(url):
    """Загружает страницу и возвращает BeautifulSoup объект"""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    response = requests.get(url, headers=headers, verify=False, timeout=10)
    response.encoding = "utf-8"
    response.raise_for_status()
    return BeautifulSoup(response.text, "html.parser")

def main():
    urls = {
        "qualification": "https://educpo.ru/dopolnitelnoe-obrazovanie/povyshenie-kvalifikacii",
        "retraining": "https://educpo.ru/dopolnitelnoe-obrazovanie/professionalnaya-perepodgotovka"
    }
    all_data = []
    
    # Парсим повышение квалификации
    try:
        soup = fetch_soup(urls["qualification"])
        qual_data = parse_qualification(soup)
        all_data.extend(qual_data)
        print(f"✅ Повышение квалификации: {len(qual_data)} программ")
    except Exception as e:
        print(f"❌ Ошибка при парсинге повышения квалификации: {e}")
    
    # Парсим переподготовку
    try:
        soup = fetch_soup(urls["retraining"])
        ret_data = parse_retraining(soup)
        all_data.extend(ret_data)
        print(f"✅ Профессиональная переподготовка: {len(ret_data)} программ")
    except Exception as e:
        print(f"❌ Ошибка при парсинге переподготовки: {e}")
    
    if not all_data:
        print("❌ Нет данных для сохранения")
        return
    
    df = pd.DataFrame(all_data)
    output_file = "all_programs.xlsx"
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"\n🎉 Сохранено всего {len(all_data)} записей в файл {output_file}")

if __name__ == "__main__":
    main()