import subprocess
import sys
import time


def run_script(script_name):
    start_time = time.time()
    try:
        subprocess.run([sys.executable, script_name], check=True)
        elapsed_time = time.time() - start_time  # Время выполнения
        print(f"{script_name} успешно выполнен за {elapsed_time:.2f} секунд.\n")
    except subprocess.CalledProcessError as e:
        print(f"Ошибка выполнения {script_name}: {e}")
        sys.exit(1)


def main():
    scripts = [
        "download.py",
        "process_files.py",
        "split_groups.py",
        "json_name_subjects.py",
        "subject_to_files_extended.py",
        "formating_otchet.py"
    ]

    for script in scripts:
        print(f"Запуск {script}...")
        run_script(script)


if __name__ == "__main__":
    main()
