import re
import requests
import pandas as pd
import io
import urllib3
import PyPDF2

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URL = "https://chipk.ru/wp-content/uploads/2026/06/2027_plan-kursov_chf-peipk.pdf"

def download_pdf(url):
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers, verify=False, timeout=30)
    resp.raise_for_status()
    return resp.content

def extract_text_with_pypdf2(pdf_bytes):
    reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def parse_courses_improved(text):
    """
    Построчный парсинг курсов.
    Каждый курс обычно содержит:
    - Номер (Ч.ХХ)
    - Название в кавычках «...»
    - Часы (цифра + слово "час")
    - Стоимость (цифра + "руб")
    """
    lines = text.split('\n')
    courses = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        # Пропускаем пустые строки и служебные заголовки
        if not line or line.startswith('Ч.') or 'курс' in line.lower():
            # Если строка начинается с "Ч." или содержит "курс", она может быть частью следующей записи
            pass

        # Ищем паттерн курса: возможен многострочный захват
        # Попробуем объединить несколько строк, если текущая строка содержит кавычку или начало номера
        if '«' in line or line.startswith('Ч.'):
            # Собираем блок до следующего явного разделителя (пустая строка или следующая "Ч.")
            block = line
            j = i + 1
            while j < len(lines) and not lines[j].strip().startswith('Ч.') and lines[j].strip() != '':
                block += ' ' + lines[j].strip()
                j += 1
            # Парсим блок
            match = re.search(r'Ч\.\d+\s*[а-я]?\s*«([^»]+)»', block, re.IGNORECASE)
            if not match:
                # Может быть без "Ч."
                match = re.search(r'«([^»]+)»', block)
            if match:
                name = match.group(1).strip()
            else:
                name = None

            # Ищем часы
            hours_match = re.search(r'(\d+)\s*(?:час(?:а|ов|)?)', block, re.IGNORECASE)
            hours = int(hours_match.group(1)) if hours_match else None

            # Ищем стоимость
            cost_match = re.search(r'Стоимость:\s*([\d\s]+)\s*(?:руб\.|рублей|руб)', block, re.IGNORECASE)
            if cost_match:
                cost_str = cost_match.group(1).replace(' ', '').strip()
                cost = int(cost_str) if cost_str.isdigit() else cost_str
            else:
                cost = None

            if name and hours and cost:
                courses.append({
                    "Наименование": name,
                    "Часы": hours,
                    "Стоимость": cost
                })
                i = j
                continue
        i += 1

    return courses

def main():
    print("Скачиваем PDF...")
    pdf_bytes = download_pdf(URL)
    print("Извлекаем текст...")
    text = extract_text_with_pypdf2(pdf_bytes)

    if len(text.strip()) < 100:
        print("PyPDF2 не дал текста, пробуем pdfplumber...")
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
        except ImportError:
            print("pdfplumber не установлен.")
            return

    print("Парсим курсы (построчный метод)...")
    courses = parse_courses_improved(text)

    if not courses:
        print("Курсы не найдены. Показываем первые 1500 символов текста:")
        print(text[:1500])
        return

    print(f"Найдено курсов: {len(courses)}")
    df = pd.DataFrame(courses)
    df.to_excel("courses_full.xlsx", index=False)
    print("✅ Данные сохранены в courses_full.xlsx")
    print("\nПервые 10 записей:")
    print(df.head(10).to_string())

if __name__ == "__main__":
    main()