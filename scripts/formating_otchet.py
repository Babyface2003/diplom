import os
import json
import pandas as pd
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

FOLDERS = [
    "1_курс", "2_курс", "3_курс",
    "4_курс", "5_курс",
    "1_курс_мага", "2_курс_мага"
]


def add_empty_rows(df, num_rows=6):
    wb = Workbook()
    ws = wb.active
    for _ in range(num_rows):
        ws.append([])
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    return wb


def add_names_columns(df):
    if "М1" not in df.columns:
        df["М1"] = ""
    if "М2" not in df.columns:
        df["М2"] = ""
    if "Примечание" not in df.columns:
        df["Примечание"] = ""
    return df


def combine_columns(df):
    cols = {"Фамилия", "Имя", "Отчество"}
    if cols.issubset(df.columns):
        for c in ["Фамилия", "Имя"]:
            df[c] = df[c].astype(str).str.strip()
        df["ФИО"] = (df["Фамилия"] + " " + df["Имя"]).str.strip()
        df.drop(["Фамилия", "Имя", "Отчество"], axis=1, inplace=True)
    return df


def process_dataframe(df):
    # Обработка первой колонки: сохраняем номер
    first_col = df.columns[0]
    if first_col.startswith("Unnamed") or first_col == df.index.name:
        if first_col != '№':
            df.rename(columns={first_col: '№'}, inplace=True)
    # Вставляем новый столбец №, только если его ещё нет
    if '№' not in df.columns:
        df.insert(0, '№', range(1, len(df) + 1))

    # Обработка ФИО и добавление столбцов
    df = combine_columns(df)
    df = add_names_columns(df)

    # Переставляем колонки в нужном порядке и вставляем пустой сдвиг
    df = df[['№', 'ФИО', 'М1', 'М2', 'Примечание']]
    df.insert(2, '', '')  # пустая колонка для сдвига ФИО на B:C
    return df


def shorten_subject(subject):
    words = subject.strip().split()
    letters = []
    for i, w in enumerate(words):
        if w:
            letters.append(w[0].upper() if i == 0 else w[0].lower())
    return "".join(letters)


def set_worksheet_formats(ws):
    # Заголовок
    ws.merge_cells("A1:H1")
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # Объединение ячеек по строкам
    row_h = 7
    for r in range(row_h, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if all(v is None or str(v).strip() == "" for v in vals):
            break
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=8)

    # Фиксированная ширина для первого столбца
    ws.column_dimensions['A'].width = 5  # здесь можно указать нужное значение
    # Автоподбор ширины остальных столбцов
    for col in range(2, ws.max_column + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for row in range(1, ws.max_row + 1):
            v = ws.cell(row=row, column=col).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[col_letter].width = max_len + 2


def combined_excel_files():
    path = "../new_data"
    os.makedirs(path, exist_ok=True)
    output = Path(path) / "combined.xlsx"
    if output.exists():
        output.unlink()

    with open("subject_to_files.json", "r", encoding="utf-8") as f:
        subject_groups = json.load(f)
    with open("group_directions.json", "r", encoding="utf-8") as f:
        group_dirs = json.load(f)
    with open("subject_to_files_extended.json", "r", encoding="utf-8") as f:
        extended_subject_data = json.load(f)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for folder in FOLDERS:
            folder_path = Path("data") / folder
            if not folder_path.exists():
                continue
            for file_path in folder_path.glob("*.xlsx"):
                df = pd.read_excel(file_path)
                df = process_dataframe(df)
                for subject, groups in subject_groups.items():
                    for group in groups:
                        short_subj = shorten_subject(subject)
                        sheet_name = f"{short_subj} {group}"
                        df.to_excel(writer, sheet_name=sheet_name, startrow=6, startcol=0, index=False)
                        ws = writer.sheets[sheet_name]
                        ws["A1"] = subject
                        ws["H3"] = group
                        ws["H4"] = group_dirs.get(group, "")
                        ws["B3"] = "Лектор"
                        ws["B4"] = "Семинарист"
                        ws["B5"] = "Лаборант"

                        subject_data = extended_subject_data.get(subject, [])
                        for entry in subject_data:
                            ttype = entry.get("type", "")
                            teacher = entry.get("teacher", "")
                            egroup = entry.get("group", "")
                            if ttype == "Лекции":
                                ws["C3"] = teacher
                            elif egroup.startswith(group):
                                if ttype == "Практические":
                                    ws["C4"] = teacher
                                elif ttype == "Лабораторные":
                                    ws["C5"] = teacher
    wb = load_workbook(output)
    for ws in wb.worksheets:
        set_worksheet_formats(ws)
    wb.save(output)
    print("Done", output)


def main():
    combined_excel_files()


if __name__ == "__main__":
    main()
