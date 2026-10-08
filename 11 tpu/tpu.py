import os
import time
import signal
import sys
from typing import List, Dict

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Конфигурация
URL = "https://dpo.tpu.ru/courses/"
OUTPUT_FILENAME = "tpu_programs.xlsx"
DELAY_BETWEEN_TABS = 2
WAIT_TIMEOUT = 10

# Настройки прокрутки
SCROLL_PAUSE_TIME = 1
MAX_SCROLL_ATTEMPTS = 30

# Для сохранения при прерывании
all_programs = []


def signal_handler(sig, frame):
    print("\n⚠️ Прерывание. Сохраняем собранные данные...")
    if all_programs:
        save_to_excel(all_programs, OUTPUT_FILENAME)
    sys.exit(0)


def save_to_excel(data: List[Dict], filename: str):
    if not data:
        return

    if os.path.exists(filename):
        try:
            existing_df = pd.read_excel(filename)
            existing_data = existing_df.to_dict('records')
        except Exception as e:
            print(f"⚠️ Не удалось загрузить существующий файл: {e}")
            existing_data = []
    else:
        existing_data = []

    combined = existing_data + data
    df = pd.DataFrame(combined)

    # Дедупликация
    df = df.drop_duplicates(
        subset=['Раздел', 'Название программы', 'Часы', 'Стоимость'],
        keep='first'
    )

    cols_order = [
        'Раздел',
        'Название программы',
        'Ссылка',
        'Часы',
        'Стоимость',
        'Комментарий'
    ]
    df = df[cols_order]
    df = df.sort_values('Раздел')
    df.to_excel(filename, index=False)
    print(f"💾 Сохранено {len(df)} записей в {filename}")


def parse_card(card, section_name: str) -> Dict:
    """Извлекает данные из одной карточки .direction"""
    hours_elem = card.find('div', class_='direction__training')

    # Часы
    hours = ""
    if hours_elem:
        short_title = hours_elem.find('span', class_='short-title')
        if short_title:
            hours = short_title.get_text(strip=True)

    # Стоимость
    price = ""
    if hours_elem:
        title_elem = hours_elem.find('span', class_='title')
        if title_elem:
            price = title_elem.get_text(strip=True)

    # Название программы + комментарий
    prog_elem = card.find('div', class_='direction__educational-program')

    comment = ""
    if prog_elem:
        comment_elem = prog_elem.find('span', class_='short-title')
        if comment_elem:
            comment = comment_elem.get_text(strip=True)

    program_name = ""
    if prog_elem:
        name_elem = prog_elem.find('h4')
        if name_elem:
            a_tag = name_elem.find('a')
            if a_tag:
                program_name = a_tag.get_text(strip=True)

    # Ссылка
    link = ""
    more_link = card.find('a', class_='more')
    if more_link and more_link.has_attr('href'):
        link = more_link['href']

    return {
        'Раздел': section_name,
        'Название программы': program_name,
        'Ссылка': link,
        'Часы': hours,
        'Стоимость': price,
        'Комментарий': comment,
    }


def scroll_to_bottom(driver, pause=SCROLL_PAUSE_TIME, max_attempts=MAX_SCROLL_ATTEMPTS):
    """Прокрутка страницы вниз до стабилизации высоты."""
    last_height = driver.execute_script("return document.body.scrollHeight")

    for _ in range(max_attempts):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause)
        new_height = driver.execute_script("return document.body.scrollHeight")

        if new_height == last_height:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(pause)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break

        last_height = new_height

    print(f"   ⬇️ Прокрутка завершена, высота страницы: {last_height}")


def scrape_with_selenium():
    global all_programs

    driver = webdriver.Firefox()
    driver.get(URL)
    time.sleep(3)

    # Закрываем возможное модальное окно
    try:
        close_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.modal__close"))
        )
        close_btn.click()
        time.sleep(1)
    except Exception:
        pass

    tab_buttons = driver.find_elements(By.CSS_SELECTOR, "button.tabs-navigation__btn")
    print(f"🔍 Найдено вкладок: {len(tab_buttons)}")

    # Один раз запрашиваем все блоки контента, чтобы проверить их количество
    tab_content_containers = driver.find_elements(By.CSS_SELECTOR, "div.js-tab")
    print(f"🔍 Найдено блоков контента (div.js-tab): {len(tab_content_containers)}")

    if len(tab_content_containers) < len(tab_buttons):
        print("⚠️ Блоков контента меньше, чем вкладок — проверьте селектор.")

    for idx, btn in enumerate(tab_buttons):
        section_name = btn.text.strip()
        if not section_name:
            continue

        print(f"\n[{idx + 1}/{len(tab_buttons)}] Обработка: {section_name}")

        # Клик по вкладке
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception as e:
            print(f"   ❌ Ошибка клика: {e}")
            continue

        # Даём анимации завершиться
        time.sleep(DELAY_BETWEEN_TABS)

        # Прокрутка страницы до конца
        scroll_to_bottom(driver)
        time.sleep(SCROLL_PAUSE_TIME)

        # Заново получаем контейнеры — они могли пересоздаться
        tab_content_containers = driver.find_elements(By.CSS_SELECTOR, "div.js-tab")
        if idx >= len(tab_content_containers):
            print(f"   ⚠️ Нет блока контента для индекса {idx}")
            continue

        # Берём контент именно этой вкладки по индексу
        content_el = tab_content_containers[idx]
        html = content_el.get_attribute('outerHTML')
        soup = BeautifulSoup(html, 'html.parser')

        cards = soup.find_all('div', class_='direction')
        print(f"   📄 Найдено программ: {len(cards)}")

        for card in cards:
            data = parse_card(card, section_name)
            if data['Название программы']:
                all_programs.append(data)

        time.sleep(DELAY_BETWEEN_TABS)

    driver.quit()


def main():
    signal.signal(signal.SIGINT, signal_handler)
    print("🚀 Начинаем парсинг dpo.tpu.ru")
    scrape_with_selenium()
    if all_programs:
        save_to_excel(all_programs, OUTPUT_FILENAME)
        print(f"\n✅ Всего собрано записей: {len(all_programs)}")
    else:
        print("❌ Данные не собраны.")


if __name__ == "__main__":
    main()