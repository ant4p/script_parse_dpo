import camelot
import pandas as pd
import requests
import urllib3
import tempfile
import os

urllib3.disable_warnings()

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

        all_dfs = []
        for table in tables:
            df = table.df.copy()
            # Добавляем колонку с номером страницы в начало
            df.insert(0, 'Страница', table.page)
            all_dfs.append(df)

        # Объединяем все таблицы в один DataFrame
        final_df = pd.concat(all_dfs, ignore_index=True, sort=False)
        
        output = "anodpo_all_in_one.xlsx"
        final_df.to_excel(output, index=False)
        print(f"✅ Сохранено {len(final_df)} строк в {output}")
        print("Столбцы: Страница + все столбцы исходных таблиц (выровнены по максимальному количеству)")

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