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
    # Дедупликация по ключевым полям
    df = df.drop_duplicates(subset=['Раздел', 'Название программы', 'Часы', 'Стоимость'], keep='first')
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
    # Часы (первый .short-title внутри .direction__training)
    hours_elem = card.find('div', class_='direction__training')
    hours = ""
    if hours_elem:
        short_title = hours_elem.find('span', class_='short-title')
        if short_title:
            hours = short_title.get_text(strip=True)

    # Стоимость (.title внутри .direction__training)
    price = ""
    if hours_elem:
        title_elem = hours_elem.find('span', class_='title')
        if title_elem:
            price = title_elem.get_text(strip=True)

    # Комментарий (.short-title внутри .direction__educational-program)
    comment = ""
    prog_elem = card.find('div', class_='direction__educational-program')
    if prog_elem:
        comment_elem = prog_elem.find('span', class_='short-title')
        if comment_elem:
            comment = comment_elem.get_text(strip=True)

    # Название программы (h4 a)
    program_name = ""
    if prog_elem:
        name_elem = prog_elem.find('h4')
        if name_elem:
            a_tag = name_elem.find('a')
            if a_tag:
                program_name = a_tag.get_text(strip=True)

    # Ссылка на программу (href из .more)
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
        'Комментарий': comment
    }

def scrape_with_selenium():
    global all_programs
    # Используем Firefox (как в примере) или Chrome – поменяйте при необходимости
    driver = webdriver.Firefox()
    driver.get(URL)
    time.sleep(3)

    # Закрываем возможное всплывающее окно (если есть)
    try:
        close_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.modal__close"))
        )
        close_btn.click()
        time.sleep(1)
    except:
        pass

    # Получаем все кнопки вкладок
    tab_buttons = driver.find_elements(By.CSS_SELECTOR, "button.tabs-navigation__btn")
    print(f"🔍 Найдено вкладок: {len(tab_buttons)}")

    for idx, btn in enumerate(tab_buttons, 1):
        section_name = btn.text.strip()
        if not section_name:
            continue

        print(f"\n[{idx}/{len(tab_buttons)}] Обработка: {section_name}")

        try:
            # Клик по вкладке через JavaScript, чтобы избежать перекрытий
            driver.execute_script("arguments[0].click();", btn)
            print(f"   Клик по вкладке {idx}")
        except Exception as e:
            print(f"   ❌ Ошибка клика: {e}")
            continue

        # Ждём появления активного контента с классом .js-tab._enter-to
        try:
            WebDriverWait(driver, WAIT_TIMEOUT).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".js-tab._enter-to"))
            )
            print(f"   ✅ Контент загрузился")
        except:
            print(f"   ⚠️ Контент не появился за {WAIT_TIMEOUT} сек")
            time.sleep(DELAY_BETWEEN_TABS)
            continue

        time.sleep(1)

        # Парсим карточки только для текущей вкладки (парсим каждую, но можно собрать все сразу)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        # Ищем активный блок контента (может быть несколько, но активный имеет класс _enter-to)
        active_content = soup.find('div', class_='js-tab', attrs={'class': lambda x: x and '_enter-to' in x.split()})
        if not active_content:
            print(f"   ⚠️ Не найден активный блок для {section_name}")
            continue

        cards = active_content.find_all('div', class_='direction')
        print(f"   📄 Найдено программ: {len(cards)}")

        for card in cards:
            program_data = parse_card(card, section_name)
            # Проверка, что данные не пустые (хотя бы название)
            if program_data['Название программы']:
                all_programs.append(program_data)

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