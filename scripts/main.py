import subprocess
import sys


def run_script(script_name):
    try:
        subprocess.run([sys.executable, script_name], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Ошибка выполнения {script_name}: {e}")
        sys.exit(1)


def main():
    scripts = [
        "download.py",
        "process_files.py",
        "split_groups.py"
    ]

    for script in scripts:
        print(f"Запуск {script}...")
        run_script(script)
        print(f'{script} успешно выполнен.\n')


if __name__ == "__main__":
    main()
