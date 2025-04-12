import pandas as pd
import json
import os
import re

data_dir = "data"
file_path = os.path.join("..", "ИТиВС_Сведения_к_составлению_расписания_Весна_2025.xlsx")

df = pd.read_excel(file_path, header=None, skiprows=4)


def main():
    def remove_parentheses(text):
        return re.sub(r"[\(\[].*?[\)\]]", "", text).strip()

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

    subject_files_extended = {}
    current_subject = None
    current_teacher = None

    for _, row in df.iterrows():
        value_b = row[1] if len(row) > 1 else None
        value_d = row[3] if len(row) > 3 else None
        value_g = row[6] if len(row) > 6 else None

        non_empty_cells = row.dropna()

        if pd.notna(value_b) and len(non_empty_cells) == 1:
            current_subject = value_b
            if current_subject not in subject_files_extended:
                subject_files_extended[current_subject] = []
            current_teacher = None
            continue

        if current_subject and pd.notna(value_b):
            str_value_b = str(value_b)
            if '[' in str_value_b and ']' in str_value_b:
                current_teacher = value_g
                type_value = "Лекции"
                group_value = str_value_b.split('[')[0].strip() + ']'
                entry = {
                    "lecture": value_b,
                    "type": type_value,
                    "teacher": current_teacher
                }
                subject_files_extended[current_subject].append(entry)
            else:
                cleaned_group = remove_parentheses(str_value_b)
                if cleaned_group in all_group_names:
                    teacher_value = value_g if pd.notna(value_g) else current_teacher
                    entry = {
                        "group": cleaned_group,
                        "lecture": value_b,
                        "type": value_d,
                        "teacher": teacher_value
                    }
                    subject_files_extended[current_subject].append(entry)

    with open("../scripts/subject_to_files_extended.json", "w", encoding="utf-8") as f:
        json.dump(subject_files_extended, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
