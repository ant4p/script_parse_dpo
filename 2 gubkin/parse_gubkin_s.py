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
URL = "https://dpo.gubkin.ru/napravleniya-obucheniya"
OUTPUT_FILENAME = "gubkin_programs.xlsx"
DELAY_BETWEEN_FILTERS = 3
WAIT_TIMEOUT = 10

# Данные будут собираться только после последнего клика
last_direction_data = []

def signal_handler(sig, frame):
    print("\n⚠️ Прерывание. Сохраняем собранные данные...")
    if last_direction_data:
        save_to_excel(last_direction_data, OUTPUT_FILENAME)
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
    df = df.drop_duplicates(subset=['Направление','Название программы', 'Стоимость'], keep='first')
    cols_order = [
        'Направление',
        'Название программы',
        'Ссылка на программу',
        'Форматы',
        'Вид обучения',
        'Дата начала',
        'Длительность',
        'Стоимость'
    ]
    df = df[cols_order]
    df = df.sort_values('Направление')
    df.to_excel(filename, index=False)
    print(f"💾 Сохранено {len(df)} записей в {filename}")

def parse_card(card):
    # Ссылка на программу
    link_tag = card.find_parent('a', class_='card_link') or card.find('a', class_='card_link')
    program_link = ""
    if link_tag and link_tag.has_attr('href'):
        href = link_tag['href']
        if href.startswith('/'):
            href = "https://dpo.gubkin.ru" + href
        program_link = href

    # Название программы
    title_elem = card.find('div', class_='card_title')
    title = title_elem.get_text(strip=True) if title_elem else ""

    # Направление (из card_label_top)
    direction_elem = card.find('div', class_='card_label_top')
    direction = direction_elem.get_text(strip=True) if direction_elem else ""

    # Форматы
    formats = []
    formats_block = card.find('div', class_='card_title_add')
    if formats_block:
        for span in formats_block.find_all('span', class_='card_title_label'):
            formats.append(span.get_text(strip=True))
    formats_str = ', '.join(formats) if formats else ""

    # Информация (вид, дата, длительность, стоимость)
    info_block = card.find('div', class_='card_info')
    study_type = ""
    start_date = ""
    duration = ""
    cost = ""

    if info_block:
        lines = []
        current = []
        for child in info_block.children:
            if child.name == 'br':
                lines.append(''.join(current).strip())
                current = []
            else:
                text = child.get_text(strip=True) if hasattr(child, 'get_text') else str(child).strip()
                if text:
                    current.append(text)
        if current:
            lines.append(''.join(current).strip())

        for line in lines:
            if not line:
                continue
            if 'повышение квалификации' in line or 'профессиональная переподготовка' in line:
                study_type = line
            elif any(month in line for month in ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']) and 'с ' in line:
                start_date = line
            elif any(key in line for key in ['недел', 'дн', 'месяц', 'час']):
                duration = line
            elif 'руб' in line:
                cost = line

    return {
        'Направление': direction,
        'Название программы': title,
        'Ссылка на программу': program_link,
        'Форматы': formats_str,
        'Вид обучения': study_type,
        'Дата начала': start_date,
        'Длительность': duration,
        'Стоимость': cost
    }

def scrape_with_selenium():
    global last_direction_data
    driver = webdriver.Firefox()
    driver.get(URL)
    time.sleep(3)

    # Закрываем возможное всплывающее окно
    try:
        close_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "svg.svg-icon--IconClose"))
        )
        close_btn.click()
        time.sleep(1)
    except:
        pass

    # Получаем все кнопки фильтров
    filter_buttons = driver.find_elements(By.CSS_SELECTOR, "div.myButton.bigbtn")
    print(f"🔍 Найдено фильтров: {len(filter_buttons)}")

    for idx, btn in enumerate(filter_buttons, 1):
        direction_name = btn.text.strip()
        if not direction_name:
            continue

        print(f"\n[{idx}/{len(filter_buttons)}] Обработка: {direction_name}")

        try:
            btn.click()
            print(f"   Клик по кнопке {idx}")
        except Exception as e:
            print(f"   ❌ Ошибка клика: {e}")
            continue

        # Ждём появления карточек
        try:
            WebDriverWait(driver, WAIT_TIMEOUT).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.card"))
            )
            print(f"   ✅ Карточки загрузились")
        except:
            print(f"   ⚠️ Карточки не появились за {WAIT_TIMEOUT} сек")
            time.sleep(DELAY_BETWEEN_FILTERS)
            continue

        time.sleep(1)

        # Если это последняя кнопка – парсим карточки
        if idx == len(filter_buttons):
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            cards = soup.find_all('div', class_='card')
            print(f"   📄 Найдено программ: {len(cards)}")
            for card in cards:
                prog = parse_card(card)
                last_direction_data.append(prog)

        time.sleep(DELAY_BETWEEN_FILTERS)

    driver.quit()

def main():
    signal.signal(signal.SIGINT, signal_handler)
    print("🚀 Начинаем парсинг dpo.gubkin.ru")
    scrape_with_selenium()
    if last_direction_data:
        save_to_excel(last_direction_data, OUTPUT_FILENAME)
        print(f"\n✅ Всего собрано записей: {len(last_direction_data)}")
    else:
        print("❌ Данные не собраны.")

if __name__ == "__main__":
    main()