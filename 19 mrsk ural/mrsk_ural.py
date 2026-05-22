import requests
import urllib3
from pathlib import Path

# Отключаем назойливые предупреждения от urllib3 о небезопасном запросе
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def download_file(url, save_path=None, verify_ssl=False):
    """
    Скачивает файл из интернета и сохраняет локально.
    
    Args:
        url (str): Ссылка на файл.
        save_path (str, optional): Путь для сохранения. По умолчанию имя из URL.
        verify_ssl (bool): Проверять SSL-сертификат. По умолчанию False (отключено).
    """
    if save_path is None:
        save_path = Path(url).name
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    
    try:
        print(f"Загрузка из: {url}")
        # verify=False отключает проверку SSL-сертификата
        with requests.get(url, stream=True, headers=headers, verify=verify_ssl) as response:
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            
            with open(save_path, 'wb') as file:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        file.write(chunk)
                        downloaded += len(chunk)
                        if total_size:
                            percent = (downloaded / total_size) * 100
                            print(f"\rПрогресс: {percent:.1f}% ({downloaded} / {total_size} байт)", end='')
            
            print(f"\nФайл сохранён: {save_path}")
            
    except requests.exceptions.RequestException as e:
        print(f"\nОшибка загрузки: {e}")

if __name__ == "__main__":
    file_url = "https://uc-mrsk-ural.ru/media/files/perechen_programm_na_2026.xlsx"
    download_file(file_url)