import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin
import urllib3
import warnings

# Отключаем только SSL-предупреждения (для самоподписанного сертификата)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def parse_mpei_programs():
    base_url = 'https://mpei.ru'
    url = base_url + '/Education/AdditionalEducation/Pages/programs.aspx'
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        response = requests.get(url, headers=headers, verify=False, timeout=30)
        response.raise_for_status()
        response.encoding = 'utf-8'
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при загрузке страницы: {e}")
        return

    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Находим все заголовки направлений
    direction_headers = soup.find_all('div', class_='groupheader')
    print(f"Найдено направлений: {len(direction_headers)}")
    
    all_programs = []
    
    for header in direction_headers:
        direction_name = header.get_text(strip=True)
        parent_li = header.find_parent('li', class_='dfwp-item')
        if not parent_li:
            print(f"Не найден родительский li для направления: {direction_name}")
            continue
        
        # Находим все таблицы внутри этого блока
        tables = parent_li.find_all('table', class_='DopProgList')
        print(f"  {direction_name}: таблиц {len(tables)}")
        
        for table in tables:
            # Получаем все строки таблицы
            rows = table.find_all('tr')
            for row in rows:
                # Пропускаем строки, которые содержат заголовки (th с классом DopProgListH)
                if row.find('th', class_='DopProgListH'):
                    continue
                
                # Берём все ячейки td
                cells = row.find_all('td')
                if len(cells) < 5:
                    continue
                
                # Код направления — первая ячейка
                code = cells[0].get_text(strip=True)
                if not code:
                    continue
                
                # Программа и ссылка — вторая ячейка
                program_tag = cells[1].find('a')
                if program_tag:
                    program_name = program_tag.get_text(strip=True)
                    program_link = urljoin(base_url, program_tag.get('href', ''))
                else:
                    program_name = cells[1].get_text(strip=True)
                    program_link = ''
                
                form = cells[2].get_text(strip=True)
                direction_field = cells[3].get_text(strip=True)
                department = cells[4].get_text(strip=True)
                
                all_programs.append({
                    'Код направления': code,
                    'Профиль / Программа': program_name,
                    'Форма обучения': form,
                    'Направление': direction_name,
                    'Кафедра': department,
                    'Ссылка на программу': program_link
                })
    
    print(f"Всего собрано записей: {len(all_programs)}")
    
    if not all_programs:
        print("Данные не найдены. Сохраняю HTML для отладки в debug_mpei.html")
        with open('debug_mpei.html', 'w', encoding='utf-8') as f:
            f.write(response.text)
        return
    
    # Сохраняем в Excel
    df = pd.DataFrame(all_programs)
    output_file = 'mpei_education_programs.xlsx'
    
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Программы', index=False)
            # Автоширина столбцов
            worksheet = writer.sheets['Программы']
            for column in worksheet.columns:
                max_length = 0
                col_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[col_letter].width = adjusted_width
        print(f"Парсинг завершён! Данные сохранены в файл: {output_file}")
    except Exception as e:
        print(f"Ошибка при сохранении Excel: {e}")

if __name__ == "__main__":
    parse_mpei_programs()