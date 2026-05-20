import camelot
import pandas as pd
import requests
import urllib3
import tempfile
import os
import warnings

# Игнорируем не критичное предупреждение о версиях библиотек
warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings("ignore", category=requests.packages.urllib3.exceptions.InsecureRequestWarning)

urllib3.disable_warnings()

def main():
    url = "https://oilgasec.ru/wp-content/uploads/Прайс-на-образовательные-услуги-2026_НОЦ-1.pdf"
    print("Загружаем PDF...")
    resp = requests.get(url, verify=False)
    resp.raise_for_status()

    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(resp.content)
        tmp_path = tmp.name

    try:
        print("Извлекаем таблицы (flavor='lattice')...")
        tables = camelot.read_pdf(tmp_path, pages='2-end', flavor='lattice')
        print(f"Найдено таблиц: {len(tables)}")

        all_dfs = []
        for table in tables:
            df = table.df.copy()
            # Сохраняем номер страницы в отдельную колонку
            df.insert(0, 'Страница', table.page)
            # Временно заменяем имена колонок на числа (0,1,2,...)
            df.columns = range(df.shape[1])
            all_dfs.append(df)

        # Объединяем все таблицы в один DataFrame (колонки будут выровнены по максимальному количеству)
        final_df = pd.concat(all_dfs, ignore_index=True, sort=False)
        
        # Теперь восстанавливаем заголовки: берем первую строку как имена колонок
        # Первая строка final_df содержит исходные заголовки (например, "№ п/п", "Наименование...")
        headers_row = final_df.iloc[0, 1:]  # Пропускаем колонку 'Страница' (индекс 0)
        # Создаем новый список имен колонок: 'Страница' + заголовки из первой строки
        new_columns = ['Страница'] + headers_row.tolist()
        final_df.columns = new_columns
        # Удаляем строку-заголовок из данных
        final_df = final_df[1:].reset_index(drop=True)

        # Дополнительная очистка: удаляем строки, где первая колонка ("№ п/п") является пустой или содержит "№ п/п" (повторные заголовки)
        if '№ п/п' in final_df.columns:
            # Удаляем строки, где в колонке "№ п/п" содержится "№ п/п" (повтор заголовка) или пустота
            mask = (final_df['№ п/п'] != '№ п/п') & (final_df['№ п/п'].notna()) & (final_df['№ п/п'] != '')
            final_df = final_df[mask].reset_index(drop=True)

        output = "oilgasec_programs.xlsx"
        final_df.to_excel(output, index=False)
        print(f"✅ Сохранено {len(final_df)} строк в {output}")
        print(f"Столбцы: {list(final_df.columns)}")

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