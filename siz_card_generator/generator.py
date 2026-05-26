from pathlib import Path
from re import sub

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


NORM_COLUMNS = [
    "Подразделение",
    "Должность",
    "Наименование СИЗ",
    "Норма выдачи",
    "Срок носки",
    "Основание",
    "Примечание",
]

EMPLOYEE_COLUMNS = [
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

FRONT_SHEET_NAME = "Лицевая сторона"
BACK_SHEET_NAME = "Оборотная сторона"
NORM_TABLE_HEADER_ROW = 22
NORM_TABLE_START_ROW = NORM_TABLE_HEADER_ROW + 1
ISSUE_TABLE_HEADER_ROW = 3
ISSUE_TABLE_NUMBER_ROW = 5
ISSUE_TABLE_START_ROW = ISSUE_TABLE_NUMBER_ROW + 1


def normalize(value):
    """Prepare text for comparison: remove extra spaces and ignore case."""
    if value is None:
        return ""
    return str(value).strip().lower()


def safe_filename(value):
    """Make a simple file name from employee name or personnel number."""
    name = str(value).strip()
    name = sub(r'[\\/:*?"<>|]+', "_", name)
    name = sub(r"\s+", "_", name)
    return name or "employee"


def split_full_name(full_name):
    """Split Russian full name into surname, name and patronymic if possible."""
    parts = str(full_name or "").strip().split()
    surname = parts[0] if len(parts) > 0 else ""
    name = parts[1] if len(parts) > 1 else ""
    patronymic = " ".join(parts[2:]) if len(parts) > 2 else ""
    return surname, name, patronymic


def read_table(path, required_columns):
    """Read an Excel sheet where the first row contains column headers."""
    workbook = load_workbook(path, data_only=True)
    sheet = workbook.active

    headers = [cell.value for cell in sheet[1]]
    missing_columns = [column for column in required_columns if column not in headers]
    if missing_columns:
        joined = ", ".join(missing_columns)
        raise ValueError(f"В файле {path.name} нет обязательных колонок: {joined}")

    rows = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(value is None for value in row):
            continue

        item = {}
        for column in required_columns:
            index = headers.index(column)
            item[column] = row[index]
        rows.append(item)

    return rows


def load_norms(norms_path):
    """Group norms by department and position."""
    rows = read_table(norms_path, NORM_COLUMNS)
    norms_by_key = {}

    for row in rows:
        department = normalize(row["Подразделение"])
        position = normalize(row["Должность"])
        key = (department, position)
        norms_by_key.setdefault(key, []).append(row)

    return norms_by_key


def style_generated_card(front_sheet, back_sheet, norms_count):
    """Apply basic formatting to the generated card."""
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    last_norm_row = max(NORM_TABLE_HEADER_ROW, NORM_TABLE_START_ROW + norms_count - 1)
    last_issue_row = max(20, ISSUE_TABLE_START_ROW + norms_count - 1)

    for row in front_sheet.iter_rows(min_row=NORM_TABLE_HEADER_ROW, max_row=last_norm_row, min_col=1, max_col=5):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    for row in back_sheet.iter_rows(min_row=ISSUE_TABLE_HEADER_ROW, max_row=last_issue_row, min_col=1, max_col=10):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="center")

    for cell in front_sheet[NORM_TABLE_HEADER_ROW]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_number in range(ISSUE_TABLE_HEADER_ROW, ISSUE_TABLE_START_ROW):
        for cell in back_sheet[row_number]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def create_card(employee, norms, template_path, output_dir):
    """Create one personal PPE card from the template."""
    workbook = load_workbook(template_path)
    front_sheet = workbook[FRONT_SHEET_NAME] if FRONT_SHEET_NAME in workbook.sheetnames else workbook.active
    back_sheet = workbook[BACK_SHEET_NAME] if BACK_SHEET_NAME in workbook.sheetnames else workbook.create_sheet(BACK_SHEET_NAME)

    full_name = employee["ФИО"]
    personnel_number = employee["Табельный номер"]
    surname, name, patronymic = split_full_name(full_name)

    front_sheet["A8"] = f"ЛИЧНАЯ КАРТОЧКА № {personnel_number}"
    front_sheet["B11"] = surname
    front_sheet["B12"] = name
    front_sheet["F12"] = patronymic
    front_sheet["B13"] = personnel_number
    front_sheet["B14"] = employee["Подразделение"]
    front_sheet["B15"] = employee["Должность"]
    front_sheet["B16"] = employee["Дата приема"]
    front_sheet["I11"] = employee["Пол"]
    front_sheet["I12"] = employee["Рост"]
    front_sheet["I14"] = employee["Размер одежды"]
    front_sheet["I15"] = employee["Размер обуви"]
    front_sheet["I16"] = employee["Размер головного убора"]
    front_sheet["I18"] = employee["Размер перчаток"]

    for index, norm in enumerate(norms, start=1):
        norm_row = NORM_TABLE_START_ROW + index - 1
        issue_row = ISSUE_TABLE_START_ROW + index - 1

        front_sheet.cell(row=norm_row, column=1, value=norm["Наименование СИЗ"])
        front_sheet.cell(row=norm_row, column=2, value=norm["Основание"])
        front_sheet.cell(row=norm_row, column=3, value=norm["Срок носки"])
        front_sheet.cell(row=norm_row, column=4, value=norm["Норма выдачи"])
        front_sheet.cell(row=norm_row, column=5, value=norm["Примечание"])

        back_sheet.cell(row=issue_row, column=1, value=norm["Наименование СИЗ"])
        for column in range(2, 11):
            back_sheet.cell(row=issue_row, column=column, value=None)

    style_generated_card(front_sheet, back_sheet, len(norms))

    output_dir.mkdir(parents=True, exist_ok=True)
    file_name = f"{safe_filename(personnel_number)}_{safe_filename(full_name)}.xlsx"
    workbook.save(output_dir / file_name)


def save_warnings(warnings, output_dir):
    """Save warnings to output/warnings.xlsx."""
    output_dir.mkdir(parents=True, exist_ok=True)
    warnings_path = output_dir / "warnings.xlsx"

    if not warnings:
        if warnings_path.exists():
            warnings_path.unlink()
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Предупреждения"
    sheet.append(["Табельный номер", "ФИО", "Подразделение", "Должность", "Причина"])

    for warning in warnings:
        sheet.append(
            [
                warning["Табельный номер"],
                warning["ФИО"],
                warning["Подразделение"],
                warning["Должность"],
                warning["Причина"],
            ]
        )

    for column in ["A", "B", "C", "D", "E"]:
        sheet.column_dimensions[column].width = 24

    workbook.save(warnings_path)


def run(norms_path, employees_path, template_path, output_dir):
    """Run the whole generation process."""
    norms_by_key = load_norms(norms_path)
    employees = read_table(employees_path, EMPLOYEE_COLUMNS)

    warnings = []
    created_count = 0

    for employee in employees:
        department = normalize(employee["Подразделение"])
        position = normalize(employee["Должность"])
        norms = norms_by_key.get((department, position))

        if not norms:
            warnings.append(
                {
                    "Табельный номер": employee["Табельный номер"],
                    "ФИО": employee["ФИО"],
                    "Подразделение": employee["Подразделение"],
                    "Должность": employee["Должность"],
                    "Причина": "Не найдены нормы для подразделения и должности",
                }
            )
            continue

        create_card(employee, norms, template_path, output_dir)
        created_count += 1

    save_warnings(warnings, output_dir)

    print(f"Готово. Создано карточек: {created_count}")
    print(f"Предупреждений: {len(warnings)}")
    print(f"Папка результата: {output_dir}")
