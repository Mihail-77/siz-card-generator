from pathlib import Path

from siz_card_generator.generator import create_single_card


BASE_DIR = Path(__file__).parent


employee = {
    "Табельный номер": "GUI-001",
    "ФИО": "Тестов Тимофей Тимофеевич",
    "Подразделение": "Склад",
    "Должность": "Кладовщик",
    "Дата приема": "2024-02-01",
    "Пол": "Учебно: мужской",
    "Рост": "Учебно: 178",
    "Размер одежды": "Учебно: 50",
    "Размер обуви": "Учебно: 42",
    "Размер головного убора": "Учебно: 58",
    "Размер перчаток": "Учебно: 9",
}


if __name__ == "__main__":
    output_path = create_single_card(
        employee=employee,
        norms_path=BASE_DIR / "data" / "norms.xlsx",
        template_path=BASE_DIR / "templates" / "card_template.xlsx",
        output_dir=BASE_DIR / "output",
    )

    if output_path is None:
        print("Карточка не создана. Проверьте output/warnings.xlsx")
    else:
        print(f"Создана карточка: {output_path}")
