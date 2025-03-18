import requests
from bs4 import BeautifulSoup
import os
from openpyxl.styles import Font
import xlwings as xw
from openpyxl import Workbook
from concurrent.futures import ThreadPoolExecutor

# Функция для входа и сохранения сессии
def login():
    session = requests.Session()
    login_url = 'https://edu.stankin.ru/login/index.php'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }
    login_page = session.get(login_url, headers=headers)
    soup = BeautifulSoup(login_page.text, 'html.parser')
    token = soup.find('input', {'name': 'logintoken'})

    if token:
        token_value = token['value']
    else:
        print("Не удалось найти логин-токен")
        return None

    data = {
        'username': 'st621233',
        'password': 'Boom1979',
        'logintoken': token_value
    }
    response = session.post(login_url, headers=headers, data=data)

    if 'login' not in response.url:
        print("Успешный вход!")
        return session
    else:
        print("Не удалось войти. Проверьте логин/пароль")
        return None


def download_excel(session, url, file_name):
    try:
        response = session.get(url)
        file_path = f'{file_name}.xls'
        with open(file_path, 'wb') as f:
            f.write(response.content)
        print(f'Файл "{file_name}" успешно скачан!')
        return file_path
    except Exception as e:
        print(f'Ошибка при скачивании "{file_name}": {e}')
        return None


def convert_xls_to_xlsx_with_formatting(xls_file_path, output_folder):
    if not os.path.exists(xls_file_path):
        print(f"Файл {xls_file_path} не найден для конвертации.")
        return None

    try:
        app = xw.App(visible=False)
        workbook = app.books.open(xls_file_path)
        for sheet in workbook.sheets:
            sheet_name = sheet.name
            xlsx_file_path = os.path.join(output_folder, f"{sheet_name}.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = sheet_name

            for row_idx, row in enumerate(sheet.used_range.value, start=1):
                for col_idx, cell_value in enumerate(row, start=1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                    if sheet.range((row_idx, col_idx)).font.bold:
                        cell.font = Font(bold=True)

            wb.save(xlsx_file_path)
            print(f"Файл '{xlsx_file_path}' успешно создан.")
    except Exception as e:
        print(f"Ошибка при конвертации: {e}")
    finally:
        workbook.close()
        app.quit()

    return output_folder


def process_course(session, course_name, course_url):
    file_path = download_excel(session, course_url, course_name)
    if file_path:
        output_folder = course_name.replace(' ', '_')
        os.makedirs(output_folder, exist_ok=True)
        convert_xls_to_xlsx_with_formatting(file_path, output_folder)
        os.remove(file_path)
        print(f"Файл {file_path} успешно удален.")


def main():
    courses = {
        '1 курс': 'https://edu.stankin.ru/pluginfile.php/518220/mod_folder/content/0/1%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1',
        '2 курс': 'https://edu.stankin.ru/pluginfile.php/518221/mod_folder/content/0/2%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1',
        '3 курс': 'https://edu.stankin.ru/pluginfile.php/518222/mod_folder/content/0/3%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1',
        '4 курс': 'https://edu.stankin.ru/pluginfile.php/518223/mod_folder/content/0/4%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1',
        '5 курс': 'https://edu.stankin.ru/pluginfile.php/518224/mod_folder/content/0/5%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1',
        '6 курс': 'https://edu.stankin.ru/pluginfile.php/518225/mod_folder/content/0/6%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1'
    }

    session = login()

    if session:
        with ThreadPoolExecutor() as executor:
            futures = [executor.submit(process_course, session, course_name, course_url) for course_name, course_url in courses.items()]
            for future in futures:
                future.result()
    else:
        print("Скачивание невозможно без входа")


if __name__ == '__main__':
    main()
