import requests
import pandas as pd
from bs4 import BeautifulSoup
import urllib3

# Отключаем предупреждения SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URL = "https://upc-centr.ru/node/26"

def get_soup(url):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    resp = requests.get(url, headers=headers, verify=False, timeout=10)
    resp.encoding = "utf-8"
    return BeautifulSoup(resp.text, "html.parser")

def clean_text(text):
    """Очищает текст от лишних пробелов и переносов строк."""
    if not text:
        return ""
    return " ".join(text.split()).strip()

def parse_programs(soup):
    table = soup.find("table")
    if not table:
        print("Таблица не найдена")
        return []

    rows = table.find_all("tr")
    programs = []
    
    # Начальная квалификация для всех программ до первого разделителя
    current_type = "Профессии рабочих"

    # Ключевые фразы для смены квалификации (только два раздела)
    type_keywords = {
        "ДОПОЛНИТЕЛЬНЫЕ ПРОФЕССИОНАЛЬНЫЕ ПРОГРАММЫ ПОВЫШЕНИЯ КВАЛИФИКАЦИИ": 
            "Дополнительные профессиональные программы повышения квалификации",
        "ДОПОЛНИТЕЛЬНЫЕ ПРОФЕССИОНАЛЬНЫЕ ПРОГРАММЫ ПРОФЕССИОНАЛЬНОЙ ПЕРЕПОДГОТОВКИ": 
            "Дополнительные профессиональные программы профессиональной переподготовки"
    }

    for row in rows:
        tds = row.find_all("td")

        # Строка-заголовок раздела: один td с colspan
        if len(tds) == 1 and tds[0].has_attr("colspan"):
            text = clean_text(tds[0].get_text())
            # Проверяем, содержит ли текст одну из ключевых фраз
            for key, type_name in type_keywords.items():
                if key in text.upper():
                    current_type = type_name
                    break
            continue  # переходим к следующей строке

        # Пропускаем строки-заголовки колонок (содержат th)
        if row.find("th"):
            continue

        # Для данных нужно минимум 9 ячеек (индексы 0..8)
        if len(tds) < 9:
            continue

        # Индексы: 0 - №, 1 - Язык, 2 - Код (пропускаем), 3 - Наименование,
        # 4 - Форма обучения, 5 - Нормативный срок, 6 - Дисциплины, 7 - Практика, 8 - Стоимость
        name_cell = tds[3]
        link = name_cell.find("a")
        name = clean_text(link.get_text()) if link else clean_text(name_cell.get_text())

        if not name:
            continue  # пустое наименование – пропускаем

        form = clean_text(tds[4].get_text())
        duration = clean_text(tds[5].get_text())
        disciplines = clean_text(tds[6].get_text())
        practice = clean_text(tds[7].get_text())
        cost = clean_text(tds[8].get_text())

        programs.append({
            "Категория": current_type,
            "Наименование": name,
            "Форма обучения": form,
            "Нормативный срок": duration,
            "Учебные дисциплины": disciplines,
            "Практика": practice,
            "Стоимость": cost
        })

    return programs

def main():
    print("Загружаем страницу...")
    soup = get_soup(URL)
    print("Парсим данные...")
    data = parse_programs(soup)

    if not data:
        print("Данные не найдены. Проверьте структуру страницы.")
        return

    print(f"Найдено программ: {len(data)}")
    df = pd.DataFrame(data)
    df.to_excel("programs_upc.xlsx", index=False)
    print("✅ Данные сохранены в programs_upc.xlsx")
    print("\nПервые 5 записей:")
    print(df.head().to_string())

if __name__ == "__main__":
    main()