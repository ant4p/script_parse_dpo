import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin
import urllib3

# Отключаем предупреждения об небезопасном SSL (для самоподписанного сертификата)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def parse_eipk_programs():
    url = "https://eipk.ru/about/learn/"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        # Добавляем verify=False, чтобы игнорировать ошибки самоподписанного сертификата
        response = requests.get(url, headers=headers, verify=False)
        response.raise_for_status()
        response.encoding = 'utf-8'
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при загрузке страницы: {e}")
        return None

    soup = BeautifulSoup(response.text, 'html.parser')
    sections = soup.find_all('div', class_='sectionW')
    
    if not sections:
        print("Не удалось найти секции на странице.")
        return None
    
    all_data = []
    
    for section in sections:
        section_header = section.find('div', class_='sectionH')
        section_name = section_header.get_text(strip=True) if section_header else "Название секции не найдено"
        
        table = section.find('table', class_='table-style-1')
        if not table:
            continue
        
        tbody = table.find('tbody')
        if not tbody:
            tbody = table
            
        rows = tbody.find_all('tr')
        
        for row in rows:
            cols = row.find_all('td')
            if len(cols) >= 4:
                program_number = cols[0].get_text(strip=True)
                
                link_tag = cols[1].find('a')
                program_name = link_tag.get_text(strip=True) if link_tag else cols[1].get_text(strip=True)
                program_url = urljoin(url, link_tag['href']) if link_tag and link_tag.has_attr('href') else ""
                
                education_type = cols[2].get_text(strip=True)
                hours = cols[3].get_text(strip=True)
                
                all_data.append({
                    'Секция': section_name,
                    '№ п/п': program_number,
                    'Наименование программы': program_name,
                    'Ссылка': program_url,
                    'Вид образования': education_type,
                    'Количество часов': hours
                })
    
    if all_data:
        df = pd.DataFrame(all_data)
        output_file = 'educational_programs.xlsx'
        try:
            df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"Данные успешно сохранены в файл '{output_file}'")
            print(f"Всего обработано записей: {len(df)}")
            print(f"Секции: {df['Секция'].unique()}")
            return df
        except Exception as e:
            print(f"Ошибка при сохранении Excel-файла: {e}")
            return None
    else:
        print("Не удалось найти данные для парсинга.")
        return None

if __name__ == "__main__":
    parse_eipk_programs()