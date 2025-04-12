import pandas as pd
import json
import os
import re

data_dir = "data"
file_path = os.path.join("..", "ИТиВС_Сведения_к_составлению_расписания_Весна_2025.xlsx")

df = pd.read_excel(file_path, header=None, skiprows=4)


def remove_parentheses(text):
    return re.sub(r"\(.*?\)", "", text).strip()


all_group_names = set()
course_folders = [
    "1_курс", "2_курс", "3_курс", "4_курс", "5_курс",
    "1_курс_мага", "2_курс_мага"
]

for folder in course_folders:
    folder_path = os.path.join(data_dir, folder)
    if os.path.exists(folder_path):
        for file in os.listdir(folder_path):
            name, _ = os.path.splitext(file)
            all_group_names.add(name.strip())

subject_files_cleaned = {}
current_subject = None

for _, row in df.iterrows():
    value_b = row[1]
    non_empty_cells = row.dropna()

    if pd.notna(value_b) and len(non_empty_cells) == 1:
        current_subject = value_b
        if current_subject not in subject_files_cleaned:
            subject_files_cleaned[current_subject] = []

    if current_subject and pd.notna(value_b):
        cleaned_group = remove_parentheses(value_b)
        if cleaned_group in all_group_names:
            subject_files_cleaned[current_subject].append(cleaned_group)

subject_files_cleaned = {k: list(set(v)) for k, v in subject_files_cleaned.items()}

output_path = "../scripts/subject_to_files.json"
with open(output_path, "w", encoding="utf-8") as f:
    json.dump(subject_files_cleaned, f, ensure_ascii=False, indent=2)
