import os
import time
import re
from urllib.parse import urljoin

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

# --- КОНФИГУРАЦИЯ ---
BASE_URL = "https://www.standart82.ru"
LOGIN_URL = f"{BASE_URL}/login"
SECTIONS = [
    {"name": "qualification-upgrade", "url": f"{BASE_URL}/qualification-upgrade"},
    {"name": "professional-retraining", "url": f"{BASE_URL}/professional-retraining"},
    {"name": "professional-education", "url": f"{BASE_URL}/professional-education"}
]

# Данные из .env
USER_EMAIL = os.getenv('USER_EMAIL')
USER_PASSWORD = os.getenv('USER_PASSWORD')

# Задержка между действиями (секунды)
DELAY = 1

def setup_driver():
    """Настраивает Firefox."""
    firefox_options = Options()
    firefox_options.add_argument("--width=1920")
    firefox_options.add_argument("--height=1080")
    # firefox_options.add_argument("--headless")
    driver = webdriver.Firefox(options=firefox_options)
    driver.implicitly_wait(10)
    return driver

def login(driver):
    print("🔐 Выполняется вход...")
    driver.get(LOGIN_URL)
    time.sleep(DELAY)

    try:
        email_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email'], input[name='email']"))
        )
        password_field = driver.find_element(By.CSS_SELECTOR, "input[type='password'], input[name='password']")

        email_field.clear()
        email_field.send_keys(USER_EMAIL)
        time.sleep(DELAY)
        password_field.clear()
        password_field.send_keys(USER_PASSWORD)
        time.sleep(DELAY)

        login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
        login_button.click()
        time.sleep(DELAY)

        WebDriverWait(driver, 15).until(EC.url_changes(LOGIN_URL))
        print("✅ Вход выполнен успешно!")
        time.sleep(DELAY)
        return True
    except Exception as e:
        print(f"❌ Ошибка при авторизации: {e}")
        return False

def collect_programs_from_section(driver, section):
    """Возвращает список записей (словарей) для Excel по одному разделу."""
    section_name = section['name']
    section_url = section['url']
    print(f"\n--- Обработка раздела: {section_name} ---")

    driver.get(section_url)
    time.sleep(DELAY)

    # Ждём появления ссылок на категории
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a[href*='/categories/']"))
        )
    except:
        print("⚠️ Ссылки на категории не найдены.")
        return []

    category_elements = driver.find_elements(By.CSS_SELECTOR, "a[href*='/categories/']")
    category_links = []
    for cat in category_elements:
        link = cat.get_attribute('href')
        if link and link not in category_links:
            category_links.append(link)

    print(f"  Найдено категорий: {len(category_links)}")
    if not category_links:
        return []

    all_rows = []

    for idx, cat_link in enumerate(category_links, 1):
        print(f"  → Категория {idx}/{len(category_links)}: {cat_link}")
        driver.get(cat_link)
        time.sleep(DELAY)  # задержка перед парсингом категории

        # Получаем HTML страницы категории
        soup = BeautifulSoup(driver.page_source, 'lxml')

        # Ищем карточки программ: каждая карточка — это блок, содержащий h3 с классом line-clamp-2
        # Поднимаемся от заголовка к родительскому контейнеру карточки
        program_cards = []
        title_tags = soup.find_all('h3', class_='line-clamp-2')
        for title_tag in title_tags:
            # Ищем ближайший div, который содержит и цену, и часы (обычно это .relative.z-10...)
            card = title_tag.find_parent('div', class_='relative')
            if not card:
                card = title_tag.find_parent('div', recursive=False)
            if card:
                program_cards.append(card)

        if not program_cards:
            print(f"    ❌ Не найдено карточек программ на странице.")
            continue

        # Извлекаем данные из каждой карточки
        for card in program_cards:
            # 1. Название
            title_tag = card.find('h3', class_='line-clamp-2')
            title = title_tag.get_text(strip=True) if title_tag else ''

            # 2. Цена (ищем любой текст с ₽)
            price = 'Цена не указана'
            price_part = card.find(string=re.compile(r'[\d\s]+₽'))
            if price_part:
                price = price_part.strip()
            else:
                # альтернативно ищем div с ценой
                price_div = card.find('div', class_='flex', string=re.compile(r'₽'))
                if price_div:
                    price = price_div.get_text(strip=True)

            # 3. Количество часов (ищем span с текстом, содержащим 'ч')
            hours_span = card.find('span', string=re.compile(r'\d+ч'))
            hours = hours_span.get_text(strip=True) if hours_span else ''

            # 4. Ссылка на программу (ищем ближайший a с href, содержащим /programs/)
            link_tag = card.find_parent('a', href=re.compile(r'/programs/'))
            if not link_tag:
                link_tag = card.find('a', href=re.compile(r'/programs/'))
            program_url = urljoin(BASE_URL, link_tag['href']) if link_tag else None

            all_rows.append({
                'Раздел': section_name,
                'Категория': cat_link.split('/')[-1],
                'Название программы': title,
                'Цена': price,
                'Часы': hours,
                'Ссылка': program_url
            })

        print(f"    Найдено программ: {len(program_cards)}")
        time.sleep(DELAY)  # задержка после обработки категории

    return all_rows

def save_to_excel(rows, filename="programs_data.xlsx"):
    if not rows:
        print("⚠️ Нет данных для сохранения.")
        return
    df = pd.DataFrame(rows)
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"✅ Данные сохранены в {filename}")

def main():
    print("🚀 Запуск парсера...")
    driver = setup_driver()
    try:
        if not login(driver):
            return

        all_data = []
        for section in SECTIONS:
            records = collect_programs_from_section(driver, section)
            all_data.extend(records)
            time.sleep(DELAY)  # задержка между разделами

        save_to_excel(all_data)
        print(f"📊 Всего собрано программ: {len(all_data)}")
    finally:
        driver.quit()
        print("👋 Браузер закрыт.")

if __name__ == "__main__":
    main()