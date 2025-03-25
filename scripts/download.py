import shutil
import requests
from bs4 import BeautifulSoup
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from concurrent.futures import ThreadPoolExecutor
import xlwings as xw
import time

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

def download_excel(session, url, file_name, output_dir):
    try:
        os.makedirs(output_dir, exist_ok=True)
        file_path = os.path.join(output_dir, f'{file_name}.xls')
        response = session.get(url)
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
        os.makedirs(output_folder, exist_ok=True)
        app = xw.App(visible=False)
        workbook = app.books.open(os.path.abspath(xls_file_path))
        for sheet in workbook.sheets:
            sheet_name = sheet.name
            xlsx_file_path = os.path.join(output_folder, f"{sheet_name}.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = sheet_name
            data = sheet.used_range.value
            if not data:
                continue
            for row_idx, row in enumerate(data, start=1):
                for col_idx, cell_value in enumerate(row, start=1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                    try:
                        if sheet.range((row_idx, col_idx)).font.bold:
                            cell.font = Font(bold=True)
                    except Exception:
                        pass
            wb.save(xlsx_file_path)
            print(f"Файл '{xlsx_file_path}' успешно создан.")
        workbook.close()
        app.quit()
    except Exception as e:
        print(f"Ошибка при конвертации: {e}")
    return output_folder

def process_course(session, course_name, course_url, output_dir):
    file_path = download_excel(session, course_url, course_name, output_dir)
    if file_path:
        course_folder = os.path.join(output_dir, course_name.replace(' ', '_'))
        convert_xls_to_xlsx_with_formatting(file_path, course_folder)
        os.remove(file_path)
        print(f"Файл {file_path} успешно удален.")

def main():
    courses = {
        '1 курс': 'https://edu.stankin.ru/pluginfile.php/518220/mod_folder/content/0/1%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1',
        '2 курс': 'https://edu.stankin.ru/pluginfile.php/518221/mod_folder/content/0/2%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1',
        '3 курс': 'https://edu.stankin.ru/pluginfile.php/518222/mod_folder/content/0/3%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1',
        '4 курс': 'https://edu.stankin.ru/pluginfile.php/518223/mod_folder/content/0/4%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1',
        '5 курс': 'https://edu.stankin.ru/pluginfile.php/518224/mod_folder/content/0/5%20%D0%BA%D1%83%D1%80%D1%81.xls?forcedownload=1'
    }
    output_dir = "data"
    if os.path.exists(output_dir):
        try:
            shutil.rmtree(output_dir)
        except Exception as e:
            print("Ошибка удаления предыдущей папки:", e)
            return
    os.makedirs(output_dir, exist_ok=True)
    session = login()
    if session:
        with ThreadPoolExecutor() as executor:
            futures = [executor.submit(process_course, session, course_name, course_url, output_dir)
                       for course_name, course_url in courses.items()]
            for future in futures:
                future.result()
    else:
        print("Скачивание невозможно без входа")

if __name__ == '__main__':
    main()
