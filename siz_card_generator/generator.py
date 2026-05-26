from pathlib import Path
from re import sub

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment


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
NORM_TABLE_HEADER_ROW = 29
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


def get_departments(norms_path):
    """Return unique departments from norms.xlsx."""
    rows = read_table(norms_path, NORM_COLUMNS)
    departments = {
        str(row["Подразделение"]).strip()
        for row in rows
        if row["Подразделение"] is not None and str(row["Подразделение"]).strip()
    }
    return sorted(departments)


def get_positions_by_department(norms_path, department):
    """Return positions available for the selected department."""
    rows = read_table(norms_path, NORM_COLUMNS)
    selected_department = normalize(department)
    positions = {
        str(row["Должность"]).strip()
        for row in rows
        if normalize(row["Подразделение"]) == selected_department
        and row["Должность"] is not None
        and str(row["Должность"]).strip()
    }
    return sorted(positions)


def find_norms_for_employee(norms_by_key, employee):
    """Find PPE norms by department and position."""
    department = normalize(employee["Подразделение"])
    position = normalize(employee["Должность"])
    return norms_by_key.get((department, position))


def build_warning(employee, reason):
    """Create one warning row for warnings.xlsx."""
    return {
        "Табельный номер": employee["Табельный номер"],
        "ФИО": employee["ФИО"],
        "Подразделение": employee["Подразделение"],
        "Должность": employee["Должность"],
        "Причина": reason,
    }


def style_generated_card(front_sheet, back_sheet, norms_count):
    """Apply basic formatting to the generated card."""
    last_norm_row = max(NORM_TABLE_HEADER_ROW, NORM_TABLE_START_ROW + norms_count - 1)
    last_issue_row = max(ISSUE_TABLE_START_ROW, ISSUE_TABLE_START_ROW + norms_count - 1)

    for row in front_sheet.iter_rows(min_row=NORM_TABLE_START_ROW, max_row=last_norm_row, min_col=1, max_col=108):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    for row in back_sheet.iter_rows(min_row=ISSUE_TABLE_START_ROW, max_row=last_issue_row, min_col=1, max_col=108):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="center")


def create_card(employee, norms, template_path, output_dir):
    """Create one personal PPE card from the template."""
    workbook = load_workbook(template_path)
    front_sheet = workbook[FRONT_SHEET_NAME] if FRONT_SHEET_NAME in workbook.sheetnames else workbook.active
    back_sheet = workbook[BACK_SHEET_NAME] if BACK_SHEET_NAME in workbook.sheetnames else workbook.create_sheet(BACK_SHEET_NAME)

    full_name = employee["ФИО"]
    personnel_number = employee["Табельный номер"]
    surname, name, patronymic = split_full_name(full_name)

    front_sheet["AF16"] = f"ЛИЧНАЯ КАРТОЧКА № {personnel_number}"
    front_sheet["L19"] = surname
    front_sheet["G20"] = name
    front_sheet["AV20"] = patronymic
    front_sheet["T21"] = personnel_number
    front_sheet["AE22"] = employee["Подразделение"]
    front_sheet["AA23"] = employee["Должность"]
    front_sheet["AE24"] = employee["Дата приема"]
    front_sheet["BW19"] = employee["Пол"]
    front_sheet["BX20"] = employee["Рост"]
    front_sheet["CA22"] = employee["Размер одежды"]
    front_sheet["BY23"] = employee["Размер обуви"]
    front_sheet["CJ24"] = employee["Размер головного убора"]
    front_sheet["CB27"] = employee["Размер перчаток"]

    for index, norm in enumerate(norms, start=1):
        norm_row = NORM_TABLE_START_ROW + index - 1
        issue_row = ISSUE_TABLE_START_ROW + index - 1

        front_sheet[f"A{norm_row}"] = norm["Наименование СИЗ"]
        front_sheet[f"AW{norm_row}"] = norm["Основание"]
        front_sheet[f"BO{norm_row}"] = norm["Срок носки"]
        front_sheet[f"CJ{norm_row}"] = norm["Норма выдачи"]

        back_sheet[f"A{issue_row}"] = norm["Наименование СИЗ"]
        for column in range(18, 109):
            back_sheet.cell(row=issue_row, column=column, value=None)

    style_generated_card(front_sheet, back_sheet, len(norms))

    output_dir.mkdir(parents=True, exist_ok=True)
    file_name = f"{safe_filename(personnel_number)}_{safe_filename(full_name)}.xlsx"
    workbook.save(output_dir / file_name)


def create_single_card(employee, norms_path, template_path, output_dir):
    """
    Create one card from employee data passed from code.

    This function is intended for a future GUI where employee data will be
    entered on the screen instead of being read from employees.xlsx.
    """
    missing_columns = [column for column in EMPLOYEE_COLUMNS if column not in employee]
    if missing_columns:
        joined = ", ".join(missing_columns)
        raise ValueError(f"Не хватает данных работника: {joined}")

    norms_by_key = load_norms(norms_path)
    norms = find_norms_for_employee(norms_by_key, employee)
    if not norms:
        reason = "Не найдены нормы для подразделения и должности"
        save_warnings([build_warning(employee, reason)], output_dir)
        return None

    create_card(employee, norms, template_path, output_dir)
    full_name = employee["ФИО"]
    personnel_number = employee["Табельный номер"]
    file_name = f"{safe_filename(personnel_number)}_{safe_filename(full_name)}.xlsx"
    return output_dir / file_name


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
        norms = find_norms_for_employee(norms_by_key, employee)

        if not norms:
            reason = "Не найдены нормы для подразделения и должности"
            warnings.append(build_warning(employee, reason))
            continue

        create_card(employee, norms, template_path, output_dir)
        created_count += 1

    save_warnings(warnings, output_dir)

    print(f"Готово. Создано карточек: {created_count}")
    print(f"Предупреждений: {len(warnings)}")
    print(f"Папка результата: {output_dir}")
