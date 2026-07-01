import requests
import pandas as pd
from bs4 import BeautifulSoup
import urllib3

# Отключаем предупреждения SSL (если сайт использует самоподписанный сертификат)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URL = "https://www.guz.ru/fakultety/institut-povysheniya-kvalifikatsii/pricesipk.php"

def get_soup(url):
    """Загружает страницу и возвращает BeautifulSoup объект."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    resp = requests.get(url, headers=headers, verify=False, timeout=10)
    resp.encoding = "utf-8"
    return BeautifulSoup(resp.text, "html.parser")

def clean_text(text):
    """Очищает текст от лишних пробелов и переносов строк."""
    if not text:
        return ""
    return " ".join(text.split()).strip()

def parse_courses(soup):
    """Парсит таблицу с курсами и возвращает список словарей."""
    # Находим таблицу с атрибутом border="1"
    table = soup.find("table", {"border": "1"})
    if not table:
        print("Таблица с курсами не найдена.")
        return []

    rows = table.find_all("tr")
    courses = []
    current_type = None  # будет меняться при встрече разделителей

    for row in rows:
        # Проверяем, является ли строка разделителем (один td с colspan)
        tds = row.find_all("td")
        if len(tds) == 1 and tds[0].has_attr("colspan"):
            text = clean_text(tds[0].get_text())
            if "Профессиональная переподготовка" in text:
                current_type = "Профессиональная переподготовка"
            elif "Повышение квалификации" in text:
                current_type = "Повышение квалификации"
            continue  # переходим к следующей строке

        # Пропускаем строку-заголовок (содержит th)
        if row.find("th"):
            continue

        # Для данных нужно 5 ячеек (индексы 0..4)
        if len(tds) < 5:
            continue

        # Извлекаем данные из ячеек
        # Ячейка 0: Наименование курса (внутри может быть ссылка <a>)
        name_cell = tds[0]
        link = name_cell.find("a")
        name = clean_text(link.get_text()) if link else clean_text(name_cell.get_text())

        if not name:
            continue  # пустое наименование – пропускаем

        hours = clean_text(tds[1].get_text())
        conditions = clean_text(tds[2].get_text())
        cost = clean_text(tds[3].get_text())
        document = clean_text(tds[4].get_text())

        # Если тип не определён (до первого разделителя), ставим "Неизвестно"
        if current_type is None:
            current_type = "Неизвестно"

        courses.append({
            "Тип программы": current_type,
            "Наименование": name,
            "Количество часов": hours,
            "Условия приема": conditions,
            "Стоимость, руб./чел.": cost,
            "Выдаваемый документ": document,
        })

    return courses

def main():
    print("Загружаем страницу...")
    soup = get_soup(URL)

    print("Парсим данные...")
    data = parse_courses(soup)

    if not data:
        print("Данные не найдены. Проверьте структуру страницы.")
        return

    print(f"Найдено курсов: {len(data)}")

    # Сохраняем в Excel
    df = pd.DataFrame(data)
    filename = "courses_guz.xlsx"
    df.to_excel(filename, index=False)
    print(f"✅ Данные сохранены в {filename}")

    # Показываем первые 5 записей для проверки
    print("\nПервые 5 курсов:")
    print(df.head().to_string())

if __name__ == "__main__":
    main()