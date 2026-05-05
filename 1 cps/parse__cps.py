import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import urllib3
import re

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://cpspec.ru/course-category/page/{}/"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
}
PAGES_TO_PARSE = 105
DELAY = 1

def extract_categories(block_soup):
    categories_text = block_soup.find(string=re.compile(r'Категории:'))
    if categories_text:
        parent = categories_text.find_parent()
        links = parent.find_all('a', href=True) if parent else []
        return ' / '.join([link.get_text(strip=True) for link in links])
    return ''

def extract_audience(block_soup):
    audience_div = block_soup.select_one('.tutor-course-details-widget-var.custom__course__widget')
    if audience_div:
        key_span = audience_div.find('span', string=re.compile(r'Аудитория:'))
        if key_span:
            value_span = audience_div.find('span', class_='tutor-course-details-widget-list')
            if value_span:
                return value_span.get_text(strip=True)
    return ''

def extract_hours(block_soup):
    hours = []
    hour_buttons = block_soup.select('.variable-item.button-variable-item')
    for btn in hour_buttons:
        title = btn.get('data-title', '')
        if title:
            hours.append(title)
    return ', '.join(hours) if hours else ''

def extract_prices(block_soup):
    # Сначала проверяем блок .flex.variation__wrap__price
    price_block = block_soup.select_one('.flex.variation__wrap__price')
    if price_block:
        old_tag = price_block.find('del')
        new_tag = price_block.find('ins')
        if old_tag and new_tag:
            old_price = old_tag.get_text(strip=True)
            new_price = new_tag.get_text(strip=True)
            old_price = re.sub(r'[^\d\s₽]', '', old_price)
            new_price = re.sub(r'[^\d\s₽]', '', new_price)
            return old_price, new_price
        else:
            amount = price_block.select_one('.woocommerce-Price-amount')
            if amount:
                price = amount.get_text(strip=True)
                price = re.sub(r'[^\d\s₽]', '', price)
                return '', price

    # Если не нашли, ищем в .woocommerce-variation-price
    var_price_block = block_soup.select_one('.woocommerce-variation-price .price')
    if var_price_block:
        amount = var_price_block.select_one('.woocommerce-Price-amount')
        if amount:
            price = amount.get_text(strip=True)
            price = re.sub(r'[^\d\s₽]', '', price)
            return '', price

    return '', ''

def parse_page(page_num):
    url = BASE_URL.format(page_num)
    print(f"🌐 Обработка: {url}")
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15, verify=False)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        course_blocks = soup.find_all('div', class_='total__custom')
        if not course_blocks:
            print(f"   ⚠️ Блоки .total__custom не найдены на странице {page_num}")
            return []

        page_data = []
        for block in course_blocks:
            title_tag = block.find('h3')
            if title_tag and title_tag.find('a'):
                title = title_tag.find('a').get_text(strip=True)
                link = title_tag.find('a')['href']
            else:
                title = link = ''

            categories = extract_categories(block)
            audience = extract_audience(block)
            hours = extract_hours(block)
            old_price, new_price = extract_prices(block)

            page_data.append({
                'Название курса': title,
                'Ссылка': link,
                'Категории': categories,
                'Аудитория': audience,
                'Количество часов (варианты)': hours,
                'Старая цена': old_price,
                'Цена со скидкой': new_price,
                'Страница': page_num
            })
        print(f"   ✅ Собрано {len(page_data)} курсов")
        return page_data
    except Exception as e:
        print(f"   ❌ Ошибка: {e}")
        return []

def main():
    all_courses = []
    for page in range(1, PAGES_TO_PARSE + 1):
        data = parse_page(page)
        all_courses.extend(data)
        time.sleep(DELAY)

    if not all_courses:
        print("Нет данных. Проверьте интернет или структуру сайта.")
        return

    df = pd.DataFrame(all_courses)
    df.to_excel('courses_data.xlsx', index=False)
    print(f"\n✅ Готово! Сохранено {len(all_courses)} курсов в файл courses_data.xlsx")

if __name__ == '__main__':
    main()