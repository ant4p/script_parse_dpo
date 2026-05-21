import requests
from bs4 import BeautifulSoup
import pandas as pd
import urllib3

# Отключаем предупреждения о самоподписанных сертификатах
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# URL = "https://akademik-pro.ru/nashi_kursy/kursy_povysheniya_kvalifikacii/"
URL = "https://akademik-pro.ru/nashi_kursy/professionalnyie_standarty.html"
HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}

def main():
    print("🚀 Загрузка страницы...")
    response = requests.get(URL, headers=HEADERS, timeout=15, verify=False)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')

    # Ищем таблицу с классом main_table
    table = soup.find('table', class_='main_table')
    if not table:
        print("❌ Таблица с классом 'main_table' не найдена")
        return

    rows = table.find_all('tr')
    data = []
    # Пропускаем заголовок (первая строка)
    for row in rows[1:]:
        cells = row.find_all('td')
        if len(cells) >= 4:
            num = cells[0].get_text(strip=True)
            name = cells[1].get_text(strip=True)
            name_prof = cells[2].get_text(strip=True)
            hours = cells[3].get_text(strip=True)
            # Очистка от лишних символов
            hours = hours.replace('—', '').strip()
            data.append([num, name, name_prof, hours])

    if not data:
        print("❌ Нет данных в таблице")
        return

    df = pd.DataFrame(data, columns=['№ п/п', 'Наименование программы', 
    'Должности', 'Кол-во часов'])
    # filename = 'akademik_pro_courses.xlsx'
    filename = 'akademik_pro_standarts.xlsx'
    df.to_excel(filename, index=False)
    print(f"✅ Сохранено {len(df)} записей в файл {filename}")

if __name__ == "__main__":
    main()