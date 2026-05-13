import os
import time
import signal
import sys
import re
from typing import List, Dict

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==================== КОНФИГУРАЦИЯ ====================
URL = "https://dpo-energo.ru/education/"
OUTPUT_FILENAME = "dpo_energo_courses.xlsx"
DELAY_BETWEEN_SECTIONS = 2
WAIT_TIMEOUT = 10

last_section_data = []

def signal_handler(sig, frame):
    print("\n⚠️ Прерывание. Сохраняем собранные данные...")
    if last_section_data:
        save_to_excel(last_section_data, OUTPUT_FILENAME)
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
    df = df.drop_duplicates(subset=['Название курса', 'Количество часов'], keep='first')
    df = df.sort_values('Название курса')
    df.to_excel(filename, index=False)
    print(f"💾 Сохранено {len(df)} записей в {filename}")

def parse_course_row(row, section_name: str = ""):
    title_div = row.find('div', class_='font-bold')
    if not title_div:
        return None
    course_name = title_div.get_text(strip=True)
    details_div = row.find('div', class_=re.compile(r'text-\[10px\]'))
    if not details_div:
        return None
    details_text = details_div.get_text(strip=True)
    hours = ""
    form = ""
    hours_match = re.search(r'(\d+)\s*ч\.', details_text)
    if hours_match:
        hours = hours_match.group(1) + ' ч.'
    form_match = re.search(r'Форма\s*-\s*([а-яА-ЯёЁ\s/]+)', details_text)
    if form_match:
        form = form_match.group(1).strip()
    return {
        'Название курса': course_name,
        'Количество часов': hours,
        'Форма обучения': form,
        'Направление/секция': section_name
    }

def get_section_name_from_button(driver, button):
    try:
        section = driver.execute_script("return arguments[0].closest('section')", button)
        if section:
            h2 = section.find_element(By.TAG_NAME, "h2")
            return h2.text.strip()
    except:
        pass
    try:
        span = button.find_element(By.CSS_SELECTOR, "span.text-xs.font-black")
        return span.text.strip()
    except:
        return "Неизвестная секция"

def scrape_with_selenium():
    global last_section_data
    options = webdriver.FirefoxOptions()
    # options.add_argument('--headless')
    driver = webdriver.Firefox(options=options)
    driver.get(URL)
    time.sleep(3)

    # Закрыть возможное всплывающее окно
    try:
        close_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[aria-label='Закрыть'], .modal-close"))
        )
        close_btn.click()
        time.sleep(1)
    except:
        pass

    # Найти все кнопки
    buttons = driver.find_elements(By.CSS_SELECTOR, "button[onclick*='toggleCourses']")
    total_buttons = len(buttons)
    print(f"🔍 Найдено секций: {total_buttons}")

    # Подготовить информацию
    buttons_info = []
    for btn in buttons:
        onclick = btn.get_attribute('onclick')
        match = re.search(r"toggleCourses\('([^']+)'", onclick)
        if match:
            block_id = match.group(1)
            section_name = get_section_name_from_button(driver, btn)
            buttons_info.append({'block_id': block_id, 'section_name': section_name})
        else:
            print(f"⚠️ Пропущена кнопка: не удалось извлечь ID")

    # Обработать каждую секцию
    for idx, info in enumerate(buttons_info, 1):
        block_id = info['block_id']
        section_name = info['section_name']
        print(f"\n[{idx}/{total_buttons}] Обработка: {section_name} (ID: {block_id})")

        # Найти актуальную кнопку
        try:
            btn = driver.find_element(By.CSS_SELECTOR, f"button[onclick*='{block_id}']")
        except Exception as e:
            print(f"   ❌ Не удалось найти кнопку: {e}")
            continue

        # Прокрутить к кнопке (центр экрана)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
        time.sleep(0.5)

        # Клик через JS, чтобы обойти перекрытие
        try:
            driver.execute_script("arguments[0].click();", btn)
            print(f"   ✅ Клик выполнен (JS)")
        except Exception as e:
            print(f"   ❌ Ошибка клика: {e}")
            continue

        # Ждать появления блока с курсами
        try:
            WebDriverWait(driver, WAIT_TIMEOUT).until(
                EC.visibility_of_element_located((By.ID, block_id))
            )
            print(f"   ✅ Блок {block_id} раскрыт")
        except Exception as e:
            print(f"   ⚠️ Блок {block_id} не появился: {e}")
            time.sleep(DELAY_BETWEEN_SECTIONS)
            continue

        time.sleep(1)

        # Парсинг
        block_element = driver.find_element(By.ID, block_id)
        block_html = block_element.get_attribute('innerHTML')
        soup_block = BeautifulSoup(block_html, 'html.parser')
        table = soup_block.find('table', class_=re.compile('border-collapse'))
        if not table:
            print(f"   ⚠️ Таблица не найдена")
            time.sleep(DELAY_BETWEEN_SECTIONS)
            continue

        rows = table.find_all('tr')
        courses_count = 0
        for row in rows:
            course_data = parse_course_row(row, section_name)
            if course_data:
                last_section_data.append(course_data)
                courses_count += 1
        print(f"   📄 Добавлено курсов: {courses_count}")

        time.sleep(DELAY_BETWEEN_SECTIONS)

    driver.quit()

def main():
    signal.signal(signal.SIGINT, signal_handler)
    print("🚀 Запуск парсера dpo-energo.ru/education")
    scrape_with_selenium()
    if last_section_data:
        save_to_excel(last_section_data, OUTPUT_FILENAME)
        print(f"\n✅ Всего собрано записей: {len(last_section_data)}")
    else:
        print("❌ Данные не собраны.")

if __name__ == "__main__":
    main()