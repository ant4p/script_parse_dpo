import requests
from bs4 import BeautifulSoup
import pandas as pd
import re


def fetch_page_content(url):
    """Загружает HTML-код страницы."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        response.encoding = 'utf-8'
        return response.text
    except requests.RequestException as e:
        print(f"Ошибка загрузки: {e}")
        return None


def parse_html_tables(html_content, max_tables=4):
    """
    Извлекает все HTML-таблицы (<table>) со страницы,
    преобразует их в DataFrame и возвращает список первых max_tables штук.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    all_tables = soup.find_all('table')
    
    if not all_tables:
        print("HTML-таблицы не найдены.")
        return []
    
    print(f"Найдено HTML-таблиц: {len(all_tables)}")
    
    dataframes = []
    for i, table in enumerate(all_tables[:max_tables]):
        try:
            # Извлекаем строки таблицы
            rows = []
            for row in table.find_all('tr'):
                cells = row.find_all(['td', 'th'])
                row_data = [cell.get_text(strip=True) for cell in cells]
                if row_data:
                    rows.append(row_data)
            
            if not rows:
                continue
            
            # Если первая строка — заголовки, используем её
            if len(rows) > 1:
                df = pd.DataFrame(rows[1:], columns=rows[0])
            else:
                df = pd.DataFrame(rows)
            
            dataframes.append(df)
            print(f"  Таблица {i+1}: {df.shape[0]} строк, {df.shape[1]} столбцов")
        except Exception as e:
            print(f"Ошибка в таблице {i+1}: {e}")
    
    return dataframes


def combine_tables(dataframes):
    """
    Объединяет несколько DataFrame в один.
    Если столбцы различаются, недостающие заполняются NaN.
    """
    if not dataframes:
        return pd.DataFrame()
    
    if len(dataframes) == 1:
        return dataframes[0]
    
    # Приводим все DataFrame к единому набору столбцов (объединение всех уникальных колонок)
    from functools import reduce
    all_columns = set()
    for df in dataframes:
        all_columns.update(df.columns)
    all_columns = sorted(list(all_columns))  # сортируем для стабильности
    
    # Добавляем недостающие столбцы в каждый DataFrame
    aligned_dfs = []
    for df in dataframes:
        for col in all_columns:
            if col not in df.columns:
                df[col] = pd.NA
        aligned_dfs.append(df[all_columns])
    
    # Объединяем вертикально
    combined = pd.concat(aligned_dfs, ignore_index=True)
    return combined


def save_combined_to_excel(df, filename="educational_programs_combined.xlsx"):
    """Сохраняет объединённую таблицу в Excel (одна вкладка)."""
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Все программы", index=False)
    print(f"Сохранено в {filename}, всего строк: {len(df)}")


def main():
    url = "https://stc-patriot.ru/prajs-list/"
    print(f"Загрузка страницы: {url}")
    html = fetch_page_content(url)
    if not html:
        return
    
    # Берём первые 4 HTML-таблицы
    tables = parse_html_tables(html, max_tables=4)
    
    if not tables:
        print("Не удалось извлечь таблицы.")
        return
    
    # Объединяем в одну
    combined_df = combine_tables(tables)
    
    # Сохраняем
    save_combined_to_excel(combined_df)


if __name__ == "__main__":
    main()