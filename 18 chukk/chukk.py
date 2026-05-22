import requests
import urllib3
from bs4 import BeautifulSoup
import pandas as pd
import re

# Отключение предупреждений о небезопасном соединении
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def parse_educational_programs(url):
    """Парсинг образовательных программ с сайта chukk.ru"""
    # Загрузка страницы с отключенной проверкой SSL
    response = requests.get(url, verify=False)
    response.encoding = 'windows-1251'  # Кодировка сайта
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Поиск контейнера с данными
    a_a_div = soup.find('div', class_='a_a')
    if not a_a_div:
        raise ValueError("Не найден контейнер с классом 'a_a'")
    
    # Поиск всех блоков с таблицами (div с классом 'text')
    text_blocks = a_a_div.find_all('div', class_='text')
    
    all_programs = []
    
    for block in text_blocks:
        # Поиск всех таблиц внутри блока
        tables = block.find_all('table')
        for table in tables:
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all('td')
                if not cells:
                    continue
                
                program = {}
                
                if len(cells) >= 3:
                    program['№ п/п'] = cells[0].get_text(strip=True) if len(cells) > 0 else ''
                    program['Код'] = cells[1].get_text(strip=True) if len(cells) > 1 else ''
                    program['Наименование'] = cells[2].get_text(strip=True) if len(cells) > 2 else ''
                    
                    if len(cells) > 5:
                        program['Теоретическое обучение (час)'] = cells[3].get_text(strip=True) if len(cells) > 3 else ''
                        program['Производственное обучение (час)'] = cells[4].get_text(strip=True) if len(cells) > 4 else ''
                        program['Общая продолжительность (час)'] = cells[5].get_text(strip=True) if len(cells) > 5 else ''
                    
                    if len(cells) > 6:
                        program['Стоимость очного обучения (руб)'] = cells[6].get_text(strip=True) if len(cells) > 6 else ''
                    if len(cells) > 7:
                        program['Стоимость электронного обучения (руб)'] = cells[7].get_text(strip=True) if len(cells) > 7 else ''
                    
                    if program['Наименование']:
                        all_programs.append(program)
    
    return all_programs

def save_to_excel(programs, filename='educational_programs.xlsx'):
    """Сохранение данных в Excel файл"""
    if not programs:
        print("Нет данных для сохранения")
        return
    
    df = pd.DataFrame(programs)
    
    column_order = ['№ п/п', 'Код', 'Наименование', 
                    'Теоретическое обучение (час)', 'Производственное обучение (час)', 
                    'Общая продолжительность (час)', 
                    'Стоимость очного обучения (руб)', 'Стоимость электронного обучения (руб)']
    
    existing_columns = [col for col in column_order if col in df.columns]
    df = df[existing_columns]
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Образовательные программы', index=False)
        
        worksheet = writer.sheets['Образовательные программы']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"Данные успешно сохранены в файл: {filename}")
    print(f"Всего программ: {len(programs)}")

if __name__ == "__main__":
    url = "https://chukk.ru/perechen-obrazovatelnykh-uslug"
    
    try:
        programs = parse_educational_programs(url)
        if programs:
            save_to_excel(programs)
        else:
            print("Не удалось извлечь данные. Возможно, структура страницы изменилась.")
    except Exception as e:
        print(f"Произошла ошибка: {e}")