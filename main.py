from pathlib import Path

from siz_card_generator.generator import run


BASE_DIR = Path(__file__).parent


if __name__ == "__main__":
    run(
        norms_path=BASE_DIR / "norms.xlsx",
        employees_path=BASE_DIR / "employees.xlsx",
        template_path=BASE_DIR / "templates" / "card_template.xlsx",
        output_dir=BASE_DIR / "output",
    )
