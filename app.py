import os
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from siz_card_generator.generator import (
    create_single_card,
    get_departments,
    get_positions_by_department,
)


BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
NORMS_PATH = DATA_DIR / "norms.xlsx"
TEMPLATE_PATH = BASE_DIR / "templates" / "card_template.xlsx"
OUTPUT_DIR = BASE_DIR / "output"
REQUESTS_DIR = BASE_DIR / "requests"

REQUEST_TITLE = "Заявка на добавление должности в нормы выдачи СИЗ"
REQUEST_FIELDS = [
    "Дата заявки",
    "Подразделение",
    "Должность",
    "Описание выполняемых работ",
    "Основание для добавления в нормы СИЗ",
    "Предполагаемые СИЗ, если известно",
    "ФИО инициатора заявки",
    "Контакт инициатора",
    "Статус рассмотрения",
    "Решение специалиста по ОТ",
    "Комментарий специалиста по ОТ",
]
REQUEST_REQUIRED_FIELDS = [
    "Подразделение",
    "Должность",
    "Описание выполняемых работ",
    "Основание для добавления в нормы СИЗ",
    "ФИО инициатора заявки",
    "Контакт инициатора",
]
REQUEST_MULTILINE_FIELDS = {
    "Описание выполняемых работ",
    "Основание для добавления в нормы СИЗ",
    "Предполагаемые СИЗ, если известно",
}
REQUEST_REVIEW_FIELDS = {
    "Статус рассмотрения",
    "Решение специалиста по ОТ",
    "Комментарий специалиста по ОТ",
}


FIELD_LABELS = [
    "Номер карточки",
    "ФИО",
    "Дата приема",
    "Пол",
    "Рост",
    "Размер одежды",
    "Размер обуви",
    "Размер головного убора",
    "Размер перчаток",
]

FIELD_TO_EMPLOYEE_COLUMN = {
    "Номер карточки": "Табельный номер",
}

REQUIRED_FIELDS = [
    "Подразделение",
    "Должность",
    "Номер карточки",
    "ФИО",
    "Дата приема",
    "Пол",
    "Рост",
    "Размер одежды",
]

NAVIGATION_KEYS = {
    "Left",
    "Right",
    "Home",
    "End",
    "Tab",
    "Return",
    "Escape",
    "Up",
    "Down",
}


def validate_request_data(request_data):
    return [
        field
        for field in REQUEST_REQUIRED_FIELDS
        if not str(request_data.get(field, "") or "").strip()
    ]


def get_request_path(request_time=None):
    timestamp = (request_time or datetime.now()).replace(microsecond=0)
    while True:
        timestamp_text = timestamp.strftime("%Y-%m-%d_%H-%M-%S")
        request_path = REQUESTS_DIR / f"Заявка_СИЗ_{timestamp_text}.xlsx"
        if not request_path.exists():
            return request_path
        timestamp += timedelta(seconds=1)


def create_request_file(request_data, request_time=None):
    missing_fields = validate_request_data(request_data)
    if missing_fields:
        raise ValueError(f"Заполните обязательное поле: {missing_fields[0]}")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Заявка"
    sheet.merge_cells("A1:B1")
    title_cell = sheet["A1"]
    title_cell.value = REQUEST_TITLE
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    title_cell.fill = PatternFill("solid", fgColor="D9EAF7")
    sheet.row_dimensions[1].height = 36

    thin_side = Side(style="thin", color="808080")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for row_number, field in enumerate(REQUEST_FIELDS, start=3):
        label_cell = sheet.cell(row=row_number, column=1, value=field)
        value = "" if field in REQUEST_REVIEW_FIELDS else request_data.get(field, "")
        value_cell = sheet.cell(row=row_number, column=2, value=str(value or "").strip())

        label_cell.font = Font(bold=True)
        label_cell.fill = PatternFill("solid", fgColor="EAF2F8")
        label_cell.alignment = Alignment(vertical="top", wrap_text=True)
        value_cell.alignment = Alignment(vertical="top", wrap_text=True)
        label_cell.border = border
        value_cell.border = border

        if field in REQUEST_MULTILINE_FIELDS or field == "Комментарий специалиста по ОТ":
            sheet.row_dimensions[row_number].height = 55
        else:
            sheet.row_dimensions[row_number].height = 28

    sheet.column_dimensions["A"].width = 42
    sheet.column_dimensions["B"].width = 72
    sheet.freeze_panes = "A3"
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_area = f"A1:B{len(REQUEST_FIELDS) + 2}"

    REQUESTS_DIR.mkdir(parents=True, exist_ok=True)
    request_path = get_request_path(request_time)
    workbook.save(request_path)
    workbook.close()
    return request_path


class SizCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор карточек СИЗ")
        self.root.resizable(False, False)

        self.department_var = tk.StringVar()
        self.position_var = tk.StringVar()
        self.field_vars = {label: tk.StringVar() for label in FIELD_LABELS}
        self.departments = []
        self.positions = []
        self.files_ready = False
        self.last_created_card_path = None

        self.build_form()
        self.files_ready = self.check_required_files()
        if self.files_ready:
            self.load_departments()
        else:
            self.create_button["state"] = "disabled"

    def build_form(self):
        frame = ttk.Frame(self.root, padding=16)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frame, text="Подразделение").grid(row=0, column=0, sticky="w", pady=4)
        self.department_combo = ttk.Combobox(
            frame,
            textvariable=self.department_var,
            width=42,
        )
        self.department_combo.grid(row=0, column=1, sticky="ew", pady=4)
        self.department_combo.bind("<<ComboboxSelected>>", self.on_department_selected)
        self.department_combo.bind("<KeyRelease>", self.on_department_typed)

        ttk.Label(frame, text="Должность").grid(row=1, column=0, sticky="w", pady=4)
        self.position_combo = ttk.Combobox(
            frame,
            textvariable=self.position_var,
            width=42,
        )
        self.position_combo.grid(row=1, column=1, sticky="ew", pady=4)
        self.position_combo.bind("<KeyRelease>", self.on_position_typed)

        for index, label in enumerate(FIELD_LABELS, start=2):
            ttk.Label(frame, text=label).grid(row=index, column=0, sticky="w", pady=4)
            if label == "Пол":
                field = ttk.Combobox(
                    frame,
                    textvariable=self.field_vars[label],
                    values=["М", "Ж"],
                    state="readonly",
                    width=42,
                )
            else:
                field = ttk.Entry(frame, textvariable=self.field_vars[label], width=45)
            field.grid(row=index, column=1, sticky="ew", pady=4)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=len(FIELD_LABELS) + 2, column=0, columnspan=2, sticky="e", pady=(12, 0))

        ttk.Button(button_frame, text="Очистить форму", command=self.clear_form).grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        self.open_card_button = ttk.Button(
            button_frame,
            text="Открыть созданную карточку",
            command=self.open_created_card,
            state="disabled",
        )
        self.open_card_button.grid(row=0, column=1, padx=(0, 8))
        self.print_card_button = ttk.Button(
            button_frame,
            text="Печать карточки СИЗ",
            command=self.print_created_card,
            state="disabled",
        )
        self.print_card_button.grid(row=0, column=2, padx=(0, 8))
        self.create_button = ttk.Button(button_frame, text="Создать карточку", command=self.create_card)
        self.create_button.grid(row=0, column=3)

        request_button_frame = ttk.Frame(frame)
        request_button_frame.grid(
            row=len(FIELD_LABELS) + 3,
            column=0,
            columnspan=2,
            sticky="e",
            pady=(8, 0),
        )
        ttk.Button(
            request_button_frame,
            text="Открыть папку заявок",
            command=self.open_requests_folder,
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(
            request_button_frame,
            text="Заявка на добавление должности",
            command=self.open_request_form,
        ).grid(row=0, column=1)

    def check_required_files(self):
        missing_files = []
        if not NORMS_PATH.exists():
            missing_files.append(f"Файл норм СИЗ: {NORMS_PATH}")
        if not TEMPLATE_PATH.exists():
            missing_files.append(f"Шаблон карточки: {TEMPLATE_PATH}")

        if missing_files:
            messagebox.showerror(
                "Не найдены обязательные файлы",
                "Программа не может создать карточку, потому что не найдены обязательные файлы:\n\n"
                + "\n".join(missing_files),
            )
            return False

        return True

    def load_departments(self):
        if not NORMS_PATH.exists():
            messagebox.showerror(
                "Не найден файл норм",
                f"Файл data/norms.xlsx не найден.\nПоложите рабочий файл норм сюда:\n{NORMS_PATH}",
            )
            self.department_combo["values"] = []
            self.position_combo["values"] = []
            return

        try:
            departments = get_departments(NORMS_PATH)
        except Exception as error:
            messagebox.showerror("Ошибка", f"Не удалось прочитать data/norms.xlsx:\n{error}")
            departments = []

        self.departments = departments
        self.department_combo["values"] = departments
        if departments:
            self.department_var.set(departments[0])
            self.update_positions(departments[0], clear_position=True)
        else:
            messagebox.showerror(
                "Нет подразделений",
                "В data/norms.xlsx не найдены подразделения. Проверьте файл норм.",
            )

    def normalize_choice(self, value):
        return str(value or "").strip().casefold()

    def filter_values(self, values, search_text):
        search = self.normalize_choice(search_text)
        if not search:
            return values
        return [value for value in values if search in self.normalize_choice(value)]

    def find_exact_value(self, values, search_text):
        search = str(search_text or "").strip()
        for value in values:
            if str(value).strip() == search:
                return value
        return None

    def update_combo_values(self, combo, values, show_list=False):
        combo["values"] = values
        if show_list and values:
            try:
                combo.after_idle(lambda: combo.event_generate("<Down>"))
            except AttributeError:
                pass

    def on_department_selected(self, _event=None):
        department = self.find_exact_value(self.departments, self.department_var.get())
        if department:
            self.department_var.set(department)
            self.update_positions(department, clear_position=True)

    def on_department_typed(self, event=None):
        if event and event.keysym in NAVIGATION_KEYS:
            return

        text = self.department_var.get()
        matches = self.filter_values(self.departments, text)
        self.update_combo_values(self.department_combo, matches, show_list=True)

        department = self.find_exact_value(self.departments, text)
        if department:
            self.update_positions(department, clear_position=False, show_empty_message=False)
        else:
            self.positions = []
            self.update_combo_values(self.position_combo, [])
            self.position_var.set("")

    def on_position_typed(self, event=None):
        if event and event.keysym in NAVIGATION_KEYS:
            return

        department = self.find_exact_value(self.departments, self.department_var.get())
        if not department:
            self.update_combo_values(self.position_combo, [])
            return

        if not self.positions:
            self.update_positions(department, clear_position=False, show_empty_message=False)
        matches = self.filter_values(self.positions, self.position_var.get())
        self.update_combo_values(self.position_combo, matches, show_list=True)

    def update_positions(self, department, clear_position=True, show_empty_message=True):
        try:
            positions = get_positions_by_department(NORMS_PATH, department)
        except Exception as error:
            messagebox.showerror("Ошибка", f"Не удалось прочитать должности:\n{error}")
            positions = []

        self.positions = positions
        self.update_combo_values(self.position_combo, positions)

        current_position = self.find_exact_value(positions, self.position_var.get())
        if clear_position:
            self.position_var.set("")
        elif current_position:
            self.position_var.set(current_position)
        else:
            self.position_var.set("")

        if department and not positions and show_empty_message:
            messagebox.showerror(
                "Нет должностей",
                f"В data/norms.xlsx нет должностей для подразделения:\n{department}",
            )

    def collect_employee(self):
        employee = {
            "Подразделение": self.department_var.get().strip(),
            "Должность": self.position_var.get().strip(),
        }

        for label, variable in self.field_vars.items():
            employee_column = FIELD_TO_EMPLOYEE_COLUMN.get(label, label)
            employee[employee_column] = variable.get().strip()

        return employee

    def validate_employee(self, employee):
        required_to_employee_column = {
            field: FIELD_TO_EMPLOYEE_COLUMN.get(field, field)
            for field in REQUIRED_FIELDS
        }
        missing = [
            field
            for field, employee_column in required_to_employee_column.items()
            if not employee.get(employee_column)
        ]
        if missing:
            messagebox.showerror("Заполните поле", f"Заполните поле: {missing[0]}")
            return False

        department = self.find_exact_value(self.departments, employee["Подразделение"])
        if not department:
            messagebox.showerror(
                "Проверьте подразделение",
                "Подразделение не найдено в нормах СИЗ. Выберите значение из списка.",
            )
            return False

        positions = get_positions_by_department(NORMS_PATH, department)
        position = self.find_exact_value(positions, employee["Должность"])
        if not position:
            messagebox.showerror(
                "Проверьте должность",
                "Должность не найдена для выбранного подразделения. Выберите значение из списка.",
            )
            return False

        employee["Подразделение"] = department
        employee["Должность"] = position
        return True

    def create_card(self):
        if not self.check_required_files():
            return

        employee = self.collect_employee()
        if not self.validate_employee(employee):
            return

        try:
            output_path = create_single_card(
                employee=employee,
                norms_path=NORMS_PATH,
                template_path=TEMPLATE_PATH,
                output_dir=OUTPUT_DIR,
            )
        except Exception as error:
            messagebox.showerror("Ошибка", f"Не удалось создать карточку. Подробности: {error}")
            return

        if output_path is None:
            messagebox.showerror(
                "Не найдены нормы",
                "Для выбранного подразделения и должности нормы СИЗ не найдены. Проверьте файл data/norms.xlsx.",
            )
            return

        self.last_created_card_path = Path(output_path)
        self.open_card_button["state"] = "normal"
        self.print_card_button["state"] = "normal"
        messagebox.showinfo("Карточка создана", f"Карточка создана:\n{output_path}")

    def open_requests_folder(self):
        REQUESTS_DIR.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(REQUESTS_DIR)
        except OSError as error:
            messagebox.showerror("Ошибка", f"Не удалось открыть папку заявок:\n{error}")

    def open_request_form(self):
        window = tk.Toplevel(self.root)
        window.title(REQUEST_TITLE)
        window.resizable(False, False)
        window.transient(self.root)

        frame = ttk.Frame(window, padding=16)
        frame.grid(row=0, column=0, sticky="nsew")

        request_vars = {
            "Дата заявки": tk.StringVar(value=date.today().isoformat()),
            "Подразделение": tk.StringVar(value=self.department_var.get().strip()),
            "Должность": tk.StringVar(value=self.position_var.get().strip()),
            "ФИО инициатора заявки": tk.StringVar(),
            "Контакт инициатора": tk.StringVar(),
        }
        text_fields = {}
        request_state = {"last_created_path": None}
        form_fields = [
            "Подразделение",
            "Должность",
            "Описание выполняемых работ",
            "Основание для добавления в нормы СИЗ",
            "Предполагаемые СИЗ, если известно",
            "ФИО инициатора заявки",
            "Контакт инициатора",
            "Дата заявки",
        ]

        for row_number, field in enumerate(form_fields):
            ttk.Label(frame, text=field).grid(
                row=row_number,
                column=0,
                sticky="nw",
                padx=(0, 12),
                pady=4,
            )
            if field in REQUEST_MULTILINE_FIELDS:
                widget = tk.Text(frame, width=58, height=4, wrap="word")
                widget.grid(row=row_number, column=1, sticky="ew", pady=4)
                text_fields[field] = widget
            else:
                widget = ttk.Entry(frame, textvariable=request_vars[field], width=60)
                widget.grid(row=row_number, column=1, sticky="ew", pady=4)

        def save_request():
            request_data = {
                field: widget.get("1.0", "end").strip()
                for field, widget in text_fields.items()
            }
            request_data.update(
                {
                    field: variable.get().strip()
                    for field, variable in request_vars.items()
                }
            )

            missing_fields = validate_request_data(request_data)
            if missing_fields:
                messagebox.showerror(
                    "Не заполнено обязательное поле",
                    f"Заполните обязательное поле: {missing_fields[0]}",
                    parent=window,
                )
                return

            try:
                request_path = create_request_file(request_data)
            except (OSError, ValueError) as error:
                messagebox.showerror(
                    "Ошибка создания заявки",
                    f"Не удалось создать заявку:\n{error}",
                    parent=window,
                )
                return

            request_state["last_created_path"] = Path(request_path)
            open_request_button["state"] = "normal"
            messagebox.showinfo(
                "Заявка создана",
                f"Заявка создана: {request_path}. "
                "Передайте файл специалисту по охране труда.",
                parent=window,
            )

        def open_created_request():
            request_path = request_state["last_created_path"]
            if not request_path or not request_path.exists():
                messagebox.showerror(
                    "Файл не найден",
                    "Файл заявки не найден. Проверьте папку requests.",
                    parent=window,
                )
                open_request_button["state"] = "disabled"
                return

            try:
                os.startfile(request_path)
            except OSError as error:
                messagebox.showerror(
                    "Ошибка",
                    f"Не удалось открыть заявку:\n{error}",
                    parent=window,
                )

        button_frame = ttk.Frame(frame)
        button_frame.grid(
            row=len(form_fields),
            column=0,
            columnspan=2,
            sticky="e",
            pady=(12, 0),
        )
        ttk.Button(button_frame, text="Закрыть", command=window.destroy).grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        open_request_button = ttk.Button(
            button_frame,
            text="Открыть заявку",
            command=open_created_request,
            state="disabled",
        )
        open_request_button.grid(
            row=0,
            column=1,
            padx=(0, 8),
        )
        ttk.Button(button_frame, text="Сохранить заявку", command=save_request).grid(
            row=0,
            column=2,
        )

        window.grab_set()
        window.focus_set()

    def clear_form(self):
        for variable in self.field_vars.values():
            variable.set("")
        self.last_created_card_path = None
        self.open_card_button["state"] = "disabled"
        self.print_card_button["state"] = "disabled"

    def open_created_card(self):
        if not self.last_created_card_path or not self.last_created_card_path.exists():
            messagebox.showerror(
                "Файл не найден",
                "Созданная карточка не найдена. Возможно, файл был удален или перемещен.",
            )
            self.open_card_button["state"] = "disabled"
            return

        try:
            os.startfile(self.last_created_card_path)
        except OSError as error:
            messagebox.showerror("Ошибка", f"Не удалось открыть карточку. Подробности: {error}")

    def print_created_card(self):
        if not self.last_created_card_path or not self.last_created_card_path.exists():
            messagebox.showerror(
                "Файл не найден",
                "Созданная карточка не найдена. Возможно, файл был удален или перемещен.",
            )
            self.print_card_button["state"] = "disabled"
            return

        try:
            os.startfile(self.last_created_card_path)
        except OSError as error:
            messagebox.showerror("Ошибка", f"Не удалось открыть карточку. Подробности: {error}")
            return

        messagebox.showinfo(
            "Печать карточки СИЗ",
            "Карточка открыта в Excel.\n\n"
            "Для печати выберите:\n"
            "1. Файл → Печать.\n"
            "2. Печатать всю книгу.\n"
            "3. Двусторонняя печать.\n"
            "4. Переворот по длинному краю.\n\n"
            "Лист 1 — лицевая сторона карточки.\n"
            "Лист 2 — оборотная сторона карточки.",
        )


def main():
    root = tk.Tk()
    SizCardApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
