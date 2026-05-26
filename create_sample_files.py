from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


BASE_DIR = Path(__file__).parent
REFERENCE_TEMPLATE_PATH = BASE_DIR / "reference" / "Карточка СИЗ.xlsx"


def style_header(sheet):
    fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def create_norms_file(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Нормы"
    sheet.append(
        [
            "Подразделение",
            "Должность",
            "Наименование СИЗ",
            "Норма выдачи",
            "Срок носки",
            "Основание",
            "Примечание",
        ]
    )
    sheet.append(
        [
            "Производственный участок",
            "Слесарь по ремонту оборудования",
            "Пример позиции СИЗ 1 (замените на реальную норму)",
            "Укажите по утвержденным нормам",
            "Укажите срок",
            "Учебное основание, заменить на утвержденные нормы",
            "Учебная строка, не является нормой",
        ]
    )
    sheet.append(
        [
            "Склад",
            "Кладовщик",
            "Пример позиции СИЗ 2 (замените на реальную норму)",
            "Укажите по утвержденным нормам",
            "Укажите срок",
            "Учебное основание, заменить на утвержденные нормы",
            "Учебная строка, не является нормой",
        ]
    )

    style_header(sheet)
    widths = [28, 30, 46, 30, 18, 42, 34]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[chr(64 + index)].width = width

    workbook.save(path)


def create_employees_file(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Работники"
    sheet.append(
        [
            "Табельный номер",
            "ФИО",
            "Подразделение",
            "Должность",
            "Дата приема",
            "Пол",
            "Рост",
            "Размер одежды",
            "Размер обуви",
            "Размер головного убора",
            "Размер перчаток",
        ]
    )
    sheet.append(
        [
            "0001",
            "Иванов Иван Иванович",
            "Производственный участок",
            "Слесарь по ремонту оборудования",
            "2024-01-15",
            "Учебно: мужской",
            "Учебно: 180",
            "Учебно: 52",
            "Учебно: 42",
            "Учебно: 58",
            "Учебно: 10",
        ]
    )
    sheet.append(
        [
            "0002",
            "Петров Петр Петрович",
            "Склад",
            "Кладовщик",
            "2023-08-01",
            "Учебно: мужской",
            "Учебно: 176",
            "Учебно: 50",
            "Учебно: 41",
            "Учебно: 57",
            "Учебно: 9",
        ]
    )
    sheet.append(
        [
            "0003",
            "Сидорова Анна Сергеевна",
            "Бухгалтерия",
            "Бухгалтер",
            "2022-05-10",
            "Учебно: женский",
            "Учебно: 168",
            "Учебно: 46",
            "Учебно: 38",
            "Учебно: 56",
            "Учебно: 7",
        ]
    )

    style_header(sheet)
    widths = [18, 30, 30, 34, 16, 18, 16, 18, 16, 24, 18]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[chr(64 + index)].width = width

    workbook.save(path)


def create_template(path):
    if not REFERENCE_TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Не найден файл-образец: {REFERENCE_TEMPLATE_PATH}")

    workbook = load_workbook(REFERENCE_TEMPLATE_PATH)
    front_sheet = workbook["стр.1"]
    back_sheet = workbook["стр.2"]

    front_sheet.title = "Лицевая сторона"
    back_sheet.title = "Оборотная сторона"

    for sheet in workbook.worksheets:
        if sheet.title not in ["Лицевая сторона", "Оборотная сторона"]:
            workbook.remove(sheet)

    setup_print(front_sheet, "A1:DD37")
    setup_print(back_sheet, "A1:DD9")
    setup_fonts(front_sheet, back_sheet)

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def setup_print(sheet, print_area):
    sheet.print_area = print_area
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.paperSize = 9
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_view.showGridLines = False


def set_cell_font(cell, size, bold=False):
    cell.font = Font(name="Times New Roman", size=size, bold=bold)


def setup_fonts(front_sheet, back_sheet):
    for sheet in [front_sheet, back_sheet]:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.has_style:
                    set_cell_font(cell, cell.font.sz or 11, cell.font.bold)

    set_cell_font(front_sheet["BS1"], 10)
    set_cell_font(front_sheet["BS2"], 10)
    set_cell_font(front_sheet["DD10"], 12)
    set_cell_font(front_sheet["A12"], 12, bold=True)
    set_cell_font(front_sheet["AW37"], 10)
    set_cell_font(front_sheet["BR37"], 10)

    for cell in ["A29", "AW29", "BO29", "CJ29"]:
        set_cell_font(front_sheet[cell], 11)
        front_sheet[cell].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

    for cell in ["A3", "R3", "AJ3", "BS3", "AJ4", "AS4", "BA4", "BI4", "BS4", "CB4", "CJ4", "CT4"]:
        set_cell_font(back_sheet[cell], 11)
        back_sheet[cell].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

    set_cell_font(back_sheet["B8"], 11)
    set_cell_font(back_sheet["B9"], 11)


def main():
    create_norms_file(BASE_DIR / "norms.xlsx")
    create_employees_file(BASE_DIR / "employees.xlsx")
    create_template(BASE_DIR / "templates" / "card_template.xlsx")
    print("Созданы файлы norms.xlsx, employees.xlsx и templates/card_template.xlsx")


if __name__ == "__main__":
    main()
