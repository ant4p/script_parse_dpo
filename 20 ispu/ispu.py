import camelot
import pandas as pd
import requests
from io import BytesIO
import os
import re

# --- 1. Конфигурация ---
pdf_urls = [
    "http://ispu.ru/sites/default/files/2025-06/ПЛАН%202026%20генерирующие%20компании.pdf",
    "http://ispu.ru/sites/default/files/2025-06/ПЛАН%202026%20промышленные%20предприятия.pdf",
    "http://ispu.ru/sites/default/files/2025-06/ПЛАН%202026%20сетевые%20компании.pdf"
]
source_names = ["Генерирующие_компании", "Промышленные_предприятия", "Сетевые_компании"]
output_file = "учебные_программы_2026.xlsx"
output_sheet = "Все_программы"

# --- 2. Вспомогательные функции ---
def safe_str(value):
    if pd.isna(value):
        return ''
    return str(value).strip()

def merge_program_rows(df):
    """Объединяет строки одной программы (многострочные наименования)."""
    if df.empty:
        return df
    df = df.fillna('').astype(str).replace('nan', '')
    new_rows = []
    current_row = None
    for idx in range(len(df)):
        first_cell = df.iloc[idx, 0].strip()
        is_new_program = bool(re.match(r'^\d+\.', first_cell))
        if is_new_program:
            if current_row is not None:
                new_rows.append(current_row)
            current_row = df.iloc[idx].tolist()
        else:
            if current_row is None:
                current_row = df.iloc[idx].tolist()
            else:
                for col in range(df.shape[1]):
                    curr_val = df.iloc[idx, col]
                    if curr_val:
                        current_row[col] = f"{current_row[col]} {curr_val}".strip()
    if current_row is not None:
        new_rows.append(current_row)
    return pd.DataFrame(new_rows)

def extract_hours_cost_duration(df):
    hours_pat = re.compile(r'(\d+(?:[.,]\d+)?)\s*(?:ак\.?\s*час|ак\.?\s*ч|час(?:ов)?|ч\.?)', re.I)
    cost_pat  = re.compile(r'(\d+(?:[.,]\d+)?)\s*(?:руб|р\.?|₽)', re.I)
    duration_pat = re.compile(r'(\d+(?:[.,]\d+)?)\s*(?:год(?:а|ов)?|г\.|мес(?:яц(?:ев)?)?|курс(?:а|ов)?)', re.I)

    def find_first(text, pattern):
        if not isinstance(text, str):
            text = safe_str(text)
        m = pattern.search(text)
        return m.group(1).replace(',', '.') if m else ''

    df['Часы'] = ''
    df['Стоимость'] = ''
    df['Срок'] = ''

    for idx in range(len(df)):
        # Объединяем все исходные колонки (кроме добавленных) в один текст
        row_text = ' '.join([safe_str(df.iloc[idx, c]) for c in range(df.shape[1] - 3)])
        df.at[idx, 'Часы'] = find_first(row_text, hours_pat)
        df.at[idx, 'Стоимость'] = find_first(row_text, cost_pat)
        df.at[idx, 'Срок'] = find_first(row_text, duration_pat)
    return df

# --- 3. Обработка одного PDF ---
def process_pdf_from_url(url, source_name):
    print(f"Начинаю обработку: {source_name}")
    try:
        response = requests.get(url, verify=False, timeout=30)
        response.raise_for_status()
        pdf_file = BytesIO(response.content)

        tables = camelot.read_pdf(
            pdf_file,
            flavor='stream',
            pages='all',
            row_tol=10,
            split_text=False,
            edge_tol=500
        )

        if not tables:
            print(f"  Предупреждение: В PDF '{source_name}' таблицы не найдены.")
            return pd.DataFrame()

        print(f"  Найдено таблиц: {len(tables)}. Выполняется объединение...")
        combined_df = pd.concat([table.df for table in tables], ignore_index=True)

        # Очистка от пустых строк
        combined_df = combined_df.fillna('')
        combined_df = combined_df.replace(r'^\s*$', '', regex=True)
        combined_df = combined_df[~combined_df.apply(lambda row: row.astype(str).str.strip().eq('').all(), axis=1)]

        # Слияние многострочных программ
        combined_df = merge_program_rows(combined_df)

        # Извлечение параметров
        combined_df = extract_hours_cost_duration(combined_df)

        # Добавляем столбец источника первым
        combined_df.insert(0, 'Источник', source_name)

        print(f"  Обработка '{source_name}' завершена. Получено строк: {len(combined_df)}")
        return combined_df
    except Exception as e:
        print(f"  Ошибка при обработке {source_name}: {e}")
        return pd.DataFrame()

# --- 4. Основной блок ---
if __name__ == "__main__":
    print("=== Запуск парсера (все данные на одном листе, поддержка многострочных программ) ===")
    all_data = pd.DataFrame()

    for url, src in zip(pdf_urls, source_names):
        df = process_pdf_from_url(url, src)
        if not df.empty:
            all_data = pd.concat([all_data, df], ignore_index=True)
        else:
            print(f"  Пропускаем источник '{src}', нет данных.")

    if not all_data.empty:
        print(f"\nВсего собрано записей: {len(all_data)}")
        print("=== Сохранение результатов в Excel (один лист) ===")
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                all_data.to_excel(writer, sheet_name=output_sheet, index=False, header=True)
                worksheet = writer.sheets[output_sheet]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = max_length + 2
                    worksheet.column_dimensions[column_letter].width = min(adjusted_width, 60)
            print(f"✅ Файл успешно сохранён: {os.path.abspath(output_file)}")
            print(f"   Данные на листе: {output_sheet}")
        except Exception as e:
            print(f"❌ Ошибка при сохранении Excel файла: {e}")
    else:
        print("❌ Нет данных для сохранения.")

    print("\n=== Работа скрипта завершена ===")