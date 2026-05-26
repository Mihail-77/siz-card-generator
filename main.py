from pathlib import Path

from siz_card_generator.generator import run


BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"


if __name__ == "__main__":
    norms_path = DATA_DIR / "norms.xlsx"
    employees_path = DATA_DIR / "employees.xlsx"

    missing_files = [path for path in [norms_path, employees_path] if not path.exists()]
    if missing_files:
        print("Не найдены рабочие файлы:")
        for path in missing_files:
            print(f"- {path}")
        print("Положите реальные данные в папку data или создайте учебные файлы командой python create_sample_files.py.")
        raise SystemExit(1)

    run(
        norms_path=norms_path,
        employees_path=employees_path,
        template_path=BASE_DIR / "templates" / "card_template.xlsx",
        output_dir=BASE_DIR / "output",
    )
