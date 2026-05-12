import camelot
import pandas as pd
import requests
import urllib3
import tempfile
import os
import re

urllib3.disable_warnings()

VALID_TYPES = {'ОР', 'ПК', 'ПП', 'ПО', 'КЦН', 'ЕПЗ'}

def clean_cell(cell):
    return str(cell).strip() if not pd.isna(cell) else ''

def normalize_number(cell):
    text = clean_cell(cell)
    if not text:
        return ''
    normalized = re.sub(r'(\d)\s+(?=\d)', r'\1', text)
    numbers = re.findall(r'\d+', normalized)
    if len(numbers) == 0:
        return ''
    if len(numbers) == 1:
        return numbers[0]
    if len(numbers) == 2 and len(numbers[0]) <= 2 and len(numbers[1]) >= 3:
        return numbers[0] + numbers[1]
    return ' / '.join(numbers)

def is_valid_type(cell):
    val = clean_cell(cell).upper().rstrip('*')
    return val if val in VALID_TYPES else ''

def main():
    url = "https://anodpo.ru/price.pdf"
    print("Загружаем PDF...")
    resp = requests.get(url, verify=False)
    resp.raise_for_status()

    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(resp.content)
        tmp_path = tmp.name

    try:
        print("Извлекаем таблицы (flavor='lattice')...")
        tables = camelot.read_pdf(tmp_path, pages='4-end', flavor='lattice')
        print(f"Найдено таблиц: {len(tables)}")

        all_rows = []
        for table in tables:
            df = table.df
            n_cols = df.shape[1]
            if n_cols > 5:
                df = df.iloc[:, :5]
            else:
                for _ in range(5 - n_cols):
                    df[f'col{len(df.columns)}'] = ''
            df.columns = ['Вид', 'Наименование', 'Часы', 'Цена', 'Примечание']
            page_num = table.page

            i = 0
            while i < len(df):
                row = df.iloc[i]
                prog_type = clean_cell(row['Вид'])
                valid_type = is_valid_type(prog_type)
                if valid_type:
                    name = clean_cell(row['Наименование'])
                    hours = clean_cell(row['Часы'])
                    price = clean_cell(row['Цена'])
                    note = clean_cell(row['Примечание'])
                    j = i + 1
                    while j < len(df) and not is_valid_type(clean_cell(df.iloc[j]['Вид'])):
                        next_name = clean_cell(df.iloc[j]['Наименование'])
                        if next_name:
                            name += ' ' + next_name
                        if not hours and clean_cell(df.iloc[j]['Часы']):
                            hours = clean_cell(df.iloc[j]['Часы'])
                        if not price and clean_cell(df.iloc[j]['Цена']):
                            price = clean_cell(df.iloc[j]['Цена'])
                        if not note and clean_cell(df.iloc[j]['Примечание']):
                            note = clean_cell(df.iloc[j]['Примечание'])
                        j += 1
                    hours = normalize_number(hours)
                    price = normalize_number(price)
                    name = re.sub(r'\s+', ' ', name).strip()
                    note = re.sub(r'\s+', ' ', note).strip()
                    all_rows.append({
                        'Вид программы': valid_type,
                        'Наименование программы': name,
                        'Кол-во часов': hours,
                        'Стоимость, руб.': price,
                        'Примечание': note,
                        'Страница': page_num
                    })
                    i = j
                else:
                    i += 1

        if not all_rows:
            print("Не найдено записей с допустимыми видами.")
            return

        df_result = pd.DataFrame(all_rows)
        df_result = df_result.drop_duplicates()
        output = "anodpo_final_with_notes.xlsx"
        df_result.to_excel(output, index=False)
        print(f"✅ Сохранено {len(df_result)} записей в {output}")
        print("\nПример первых 10 записей:")
        print(df_result.head(10).to_string())

    except Exception as e:
        print(f"Ошибка: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            os.unlink(tmp_path)
        except PermissionError:
            pass

if __name__ == "__main__":
    main()