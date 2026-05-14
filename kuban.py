import fitz  # PyMuPDF
import pytesseract
import pandas as pd
import requests
import io
import re
import os
from PIL import Image
import numpy as np
import cv2

# Если Tesseract не в PATH, укажите полный путь (пример для Windows):
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Настройки
PDF_URL = "https://eipk.ru/Прайс%202026.pdf"
OUTPUT_EXCEL = "eipk_prices_advanced.xlsx"
DPI = 300                     # высокое разрешение для чёткости
EXPECTED_COLS = 9             # ожидаемое количество столбцов
PSM_MODE = 6                  # 6 = единый блок текста (хорошо для таблиц)
LANG = "rus+eng"              # языки
# Whitelist: разрешённые символы (русские и английские буквы, цифры, пробел, точка, запятая, косая черта, дефис, скобки)
WHITELIST = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюяA-Za-z0-9 .,/()+-"

def download_pdf(url):
    print("Загрузка PDF...")
    resp = requests.get(url, verify=False)
    resp.raise_for_status()
    return resp.content

def pdf_to_images(pdf_bytes, dpi=DPI):
    """Конвертирует PDF в список PIL Image с высоким DPI."""
    print(f"Конвертация PDF в изображения (DPI={dpi})...")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page_num in range(len(doc)):
        print(f"  Страница {page_num+1} / {len(doc)} -> изображение")
        page = doc.load_page(page_num)
        # Рендеринг с высоким DPI
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))
        images.append(img)
    doc.close()
    return images

def advanced_preprocess(image):
    """
    Продвинутая подготовка изображения для OCR:
    - CLAHE для выравнивания контраста
    - Удаление шума (fastNlMeansDenoising)
    - Адаптивная бинаризация
    - Морфологическое закрытие для соединения разорванных символов
    """
    # Convert PIL to OpenCV grayscale
    cv_img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)
    
    # 1. CLAHE (Adaptive Histogram Equalization)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    cv_img_eq = clahe.apply(cv_img)
    
    # 2. Denoising (Non-local Means)
    cv_img_den = cv2.fastNlMeansDenoising(cv_img_eq, h=20)
    
    # 3. Adaptive threshold (binary inverse)
    binary = cv2.adaptiveThreshold(cv_img_den, 255,
                                   cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV, 11, 2)
    
    # 4. Morphological closing to connect broken characters
    kernel = np.ones((2,2), np.uint8)
    morphed = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel, iterations=1)
    
    # Convert back to PIL
    return Image.fromarray(morphed)

def ocr_image(image):
    """Распознаёт текст с оптимальными настройками Tesseract."""
    custom_config = f'--oem 3 --psm {PSM_MODE} -l {LANG} -c tessedit_char_whitelist="{WHITELIST}"'
    text = pytesseract.image_to_string(image, config=custom_config)
    return text

def extract_table_from_text(text, expected_cols=EXPECTED_COLS):
    """
    Извлекает строки таблицы, разделённые двумя и более пробелами.
    Возвращает список строк (каждая строка — список колонок).
    """
    lines = text.split("\n")
    table_rows = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # Разбиваем по 2+ пробелам
        cols = re.split(r'\s{2,}', line)
        # Отбрасываем слишком короткие строки (скорее всего, мусор)
        if len(cols) < 2:
            continue
        # Обрезаем до ожидаемого количества колонок (иногда бывает лишний мусор)
        if len(cols) > expected_cols:
            cols = cols[:expected_cols]
        table_rows.append(cols)
    return table_rows

def main():
    pdf_bytes = download_pdf(PDF_URL)
    images = pdf_to_images(pdf_bytes, dpi=DPI)
    
    all_rows = []  # каждая строка: [страница, колонка1, колонка2, ...]
    
    for idx, img in enumerate(images, start=1):
        print(f"\nОбработка страницы {idx}/{len(images)}...")
        # Предобработка изображения
        processed = advanced_preprocess(img)
        # OCR
        text = ocr_image(processed)
        # Извлечение таблицы
        rows = extract_table_from_text(text, expected_cols=EXPECTED_COLS)
        if not rows:
            print(f"  На странице {idx} не найдено строк таблицы.")
            # Для отладки: сохранить распознанный текст и обработанное изображение
            # with open(f"page_{idx}_text.txt", "w", encoding="utf-8") as f:
            #     f.write(text)
            # processed.save(f"page_{idx}_processed.png")
        else:
            for row in rows:
                all_rows.append([idx] + row)
    
    if not all_rows:
        print("Таблица не найдена. Попробуйте уменьшить DPI, изменить PSM или whitelist.")
        return
    
    # Выравниваем количество колонок (заполняем пустыми строками)
    max_cols = max(len(row) for row in all_rows)
    for row in all_rows:
        while len(row) < max_cols:
            row.append("")
    
    # Создаём DataFrame
    df = pd.DataFrame(all_rows)
    # Можно переименовать столбцы, если известны названия:
    # df.columns = ['Страница'] + [f'Колонка_{i+1}' for i in range(EXPECTED_COLS)]
    
    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"\n✅ Готово! Сохранено {len(df)} строк в {OUTPUT_EXCEL}")
    print("Первые 5 строк:")
    print(df.head())

if __name__ == "__main__":
    main()