import requests
from bs4 import BeautifulSoup
import re
import time
import pandas as pd
from urllib3.exceptions import InsecureRequestWarning
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from requests.exceptions import ReadTimeout

# Отключаем предупреждения о небезопасном соединении
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}
BASE_URL = 'https://az30.ru/shop/'

session = requests.Session()
session.verify = False
session.headers.update(HEADERS)

# Повторные попытки при ошибках
retry_strategy = Retry(
    total=3,
    backoff_factor=1,
    status_forcelist=[500, 502, 503, 504],
    allowed_methods=["HEAD", "GET", "OPTIONS"]
)
adapter = HTTPAdapter(max_retries=retry_strategy)
session.mount("http://", adapter)
session.mount("https://", adapter)

def get_soup(url):
    try:
        response = session.get(url, timeout=30)
        response.raise_for_status()
        response.encoding = 'utf-8'
        return BeautifulSoup(response.text, 'html.parser')
    except ReadTimeout:
        print(f'Таймаут при загрузке {url}, пропускаем')
        return None
    except Exception as e:
        print(f'Ошибка при загрузке {url}: {e}')
        return None

def clean_category_name(name):
    return re.sub(r'\s*\(\d+\)\s*$', '', name).strip()

def extract_hours(title):
    match = re.search(r'\(([^)]+)\)', title)
    if match:
        hours = match.group(1).strip()
        clean_title = re.sub(r'\s*\([^)]+\)\s*$', '', title).strip()
        return clean_title, hours
    return title, ''

def parse_categories(soup):
    categories = []
    if not soup:
        return categories
    product_cats = soup.find_all('li', class_='product-category')
    for cat in product_cats:
        a_tag = cat.find('a')
        if not a_tag:
            continue
        href = a_tag.get('href')
        if not href:
            continue
        h2 = cat.find('h2', class_='woocommerce-loop-category__title')
        if h2:
            raw_name = h2.get_text(strip=True)
            name = clean_category_name(raw_name)
        else:
            name = 'Без названия'
        categories.append({'name': name, 'url': href})
    return categories

def parse_prices_from_table(product_soup):
    """
    Возвращает (common_prices, detailed_prices)
    common_prices: 4 ключа (индивидуально, очный курс, очно-дистанционный курс, дистанционный курс)
    detailed_prices: 8 ключей (подготовка_индивидуально, подготовка_очный, подготовка_очно-дистанционный,
                      подготовка_дистанционный, повышение_очный, повышение_очно-дистанционный,
                      повышение_дистанционный, повышение_индивидуально)
    """
    common_prices = {
        'индивидуально': None,
        'очный курс': None,
        'очно-дистанционный курс': None,
        'дистанционный курс': None
    }
    detailed_prices = {
        'подготовка_индивидуально': None,
        'подготовка_очный': None,
        'подготовка_очно-дистанционный': None,
        'подготовка_дистанционный': None,
        'повышение_очный': None,
        'повышение_очно-дистанционный': None,
        'повышение_дистанционный': None,
        'повышение_индивидуально': None
    }

    table = product_soup.find('table', class_='az30-variations-list')
    if not table:
        return common_prices, detailed_prices

    # Проходим по всем строкам таблицы
    rows = table.find_all('tr')
    for row in rows:
        cells = row.find_all('td')
        if len(cells) < 2:
            continue
        type_text = cells[0].get_text(strip=True).lower()
        price_text = cells[1].get_text(strip=True)
        price_match = re.search(r'([\d\s,]+)', price_text.replace('₽', ''))
        price = None
        if price_match:
            price_str = price_match.group(1).replace(' ', '').replace(',', '.')
            try:
                price = float(price_str)
            except:
                price = None

        # 1. Заполняем common_prices (первые попавшиеся значения)
        if common_prices['индивидуально'] is None and 'индивидуальн' in type_text:
            common_prices['индивидуально'] = price
        if common_prices['очный курс'] is None and 'очный курс' in type_text:
            common_prices['очный курс'] = price
        if common_prices['очно-дистанционный курс'] is None and 'очно-дистанцион' in type_text:
            common_prices['очно-дистанционный курс'] = price
        if common_prices['дистанционный курс'] is None and 'дистанционн' in type_text:
            common_prices['дистанционный курс'] = price

        # 2. Заполняем detailed_prices по шаблонам
        if 'подготовка/переподготовка индивидуально' in type_text:
            detailed_prices['подготовка_индивидуально'] = price
        elif 'подготовка/переподготовка очно' in type_text:
            detailed_prices['подготовка_очный'] = price
        elif 'подготовка/переподготовка очно-дистанционно' in type_text:
            detailed_prices['подготовка_очно-дистанционный'] = price
        elif 'подготовка/переподготовка дистанционно' in type_text:
            detailed_prices['подготовка_дистанционный'] = price
        elif 'повышение квалификации очно' in type_text:
            detailed_prices['повышение_очный'] = price
        elif 'повышение квалификации очно-дистанционно' in type_text:
            detailed_prices['повышение_очно-дистанционный'] = price
        elif 'повышение квалификации дистанционно' in type_text:
            detailed_prices['повышение_дистанционный'] = price
        elif 'повышение квалификации индивидуально' in type_text:
            detailed_prices['повышение_индивидуально'] = price

    return common_prices, detailed_prices

def parse_product_card(product_li, category_name):
    a_tag = product_li.find('a', class_='woocommerce-LoopProduct-link')
    if not a_tag:
        return None
    product_url = a_tag.get('href')
    h2 = a_tag.find('h2', class_='woocommerce-loop-product__title')
    if not h2:
        return None
    full_title = h2.get_text(strip=True)
    clean_title, hours = extract_hours(full_title)

    common_prices, detailed_prices = parse_prices_from_table(product_li)

    # Диапазон цен из span.price
    price_span = product_li.find('span', class_='price')
    min_price = None
    max_price = None
    if price_span:
        span_text = price_span.get_text(strip=True)
        parts = re.split(r'[–-]', span_text)
        if len(parts) >= 2:
            price_from_str = parts[0].strip()
            price_to_str = parts[1].strip()
            num_match = re.search(r'([\d\s,]+\.?\d*)', price_from_str.replace('₽', ''))
            if num_match:
                try:
                    min_price = float(num_match.group(1).replace(' ', '').replace(',', '.'))
                except:
                    pass
            num_match2 = re.search(r'([\d\s,]+\.?\d*)', price_to_str.replace('₽', ''))
            if num_match2:
                try:
                    max_price = float(num_match2.group(1).replace(' ', '').replace(',', '.'))
                except:
                    pass
        else:
            num_match = re.search(r'([\d\s,]+\.?\d*)', span_text.replace('₽', ''))
            if num_match:
                try:
                    price_val = float(num_match.group(1).replace(' ', '').replace(',', '.'))
                    min_price = max_price = price_val
                except:
                    pass

    if min_price is None:
        all_vals = list(common_prices.values()) + list(detailed_prices.values())
        valid = [v for v in all_vals if v is not None]
        if valid:
            min_price = min(valid)
            max_price = max(valid)

    return {
        'Раздел': category_name,
        'Наименование программы': clean_title,
        'количество часов': hours,
        'сумма от': min_price,
        'сумма до': max_price,
        'индивидуально': common_prices['индивидуально'],
        'очный курс': common_prices['очный курс'],
        'очно-дистанционный курс': common_prices['очно-дистанционный курс'],
        'дистанционный курс': common_prices['дистанционный курс'],
        'ссылка на программу': product_url,
        # Новые 8 колонок
        'Подготовка/переподготовка индивидуально': detailed_prices['подготовка_индивидуально'],
        'Подготовка/переподготовка очно': detailed_prices['подготовка_очный'],
        'Подготовка/переподготовка очно-дистанционно': detailed_prices['подготовка_очно-дистанционный'],
        'Подготовка/переподготовка дистанционно': detailed_prices['подготовка_дистанционный'],
        'Повышение квалификации очно': detailed_prices['повышение_очный'],
        'Повышение квалификации очно-дистанционно': detailed_prices['повышение_очно-дистанционный'],
        'Повышение квалификации дистанционно': detailed_prices['повышение_дистанционный'],
        'Повышение квалификации индивидуально': detailed_prices['повышение_индивидуально']
    }

def parse_category_products(category_url, category_name):
    all_products = []
    page = 1
    while True:
        if page == 1:
            url = category_url
        else:
            base_url = category_url.rstrip('/')
            url = f"{base_url}/page/{page}/"
        print(f"    Загрузка страницы {page}: {url}")
        soup = get_soup(url)
        if not soup:
            print(f"    Не удалось загрузить страницу {page}, прекращаем")
            break
        products = soup.find_all('li', class_='product')
        if not products:
            print(f"    На странице {page} нет товаров, завершаем категорию")
            break
        print(f"    Найдено товаров: {len(products)}")
        for prod in products:
            data = parse_product_card(prod, category_name)
            if data:
                all_products.append(data)
        pagination = soup.find('ul', class_='page-numbers')
        if pagination:
            next_link = pagination.find('a', class_='next')
            if not next_link:
                print(f"    Следующая страница не найдена, завершаем категорию")
                break
        else:
            break
        page += 1
        time.sleep(1)
    print(f"  Всего собрано товаров в категории: {len(all_products)}")
    return all_products

def main():
    print("Загрузка главной страницы...")
    main_soup = get_soup(BASE_URL)
    if not main_soup:
        print("Не удалось загрузить главную страницу")
        return

    categories = parse_categories(main_soup)
    print(f"Найдено категорий: {len(categories)}")

    all_data = []
    for idx, cat in enumerate(categories, 1):
        print(f"\n[{idx}/{len(categories)}] Обработка категории: {cat['name']}")
        products = parse_category_products(cat['url'], cat['name'])
        all_data.extend(products)
        time.sleep(1)

    df = pd.DataFrame(all_data)
    columns_order = [
        'Раздел', 'Наименование программы', 'количество часов', 'сумма от', 'сумма до',
        'индивидуально', 'очный курс', 'очно-дистанционный курс', 'дистанционный курс',
        'Подготовка/переподготовка индивидуально',
        'Подготовка/переподготовка очно',
        'Подготовка/переподготовка очно-дистанционно',
        'Подготовка/переподготовка дистанционно',
        'Повышение квалификации очно',
        'Повышение квалификации очно-дистанционно',
        'Повышение квалификации дистанционно',
        'Повышение квалификации индивидуально',
        'ссылка на программу'
    ]
    df = df[[col for col in columns_order if col in df.columns]]

    output_file = 'programs_az30.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"\nГотово! Сохранено {len(df)} записей в файл {output_file}")

if __name__ == '__main__':
    main()