import requests
import pdfplumber
import pandas as pd
from io import BytesIO
import urllib3

# Отключаем предупреждения SSL (чтобы не засоряли вывод)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://uk57.ru/wp-content/uploads/2023/04/программа-обучения-1-1.pdf"

# Скачиваем с игнорированием проверки SSL
try:
    response = requests.get(url, verify=False, timeout=30)
    response.raise_for_status()
except requests.exceptions.SSLError as e:
    print(f"Ошибка SSL: {e}")
    print("Попробуйте другой способ: скачайте файл вручную и укажите локальный путь.")
    print("Или используйте verify=False как показано выше.")
    exit(1)

# Далее обработка, как ранее
with pdfplumber.open(BytesIO(response.content)) as pdf:
    all_tables = []
    for page_num, page in enumerate(pdf.pages, start=1):
        tables = page.extract_tables()
        for table in tables:
            if table and len(table) > 1:  # проверка, что есть данные
                # Берём первую строку как заголовки
                headers = table[0]
                data = table[1:]
                # Удаляем строки, где все значения None
                data = [row for row in data if any(cell is not None and cell.strip() for cell in row)]
                if data:
                    df = pd.DataFrame(data, columns=headers)
                    df['Page'] = page_num
                    all_tables.append(df)
    
    if all_tables:
        final_df = pd.concat(all_tables, ignore_index=True)
        final_df.to_excel("output.xlsx", index=False, engine='openpyxl')
        print("Файл output.xlsx успешно создан.")
    else:
        print("Таблицы в PDF не найдены. Попробуйте другие настройки извлечения.")