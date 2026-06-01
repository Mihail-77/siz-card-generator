import os
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

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

        ttk.Button(button_frame, text="Открыть папку output", command=self.open_output_folder).grid(
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
        self.create_button = ttk.Button(button_frame, text="Создать карточку", command=self.create_card)
        self.create_button.grid(row=0, column=2)

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
        messagebox.showinfo("Карточка создана", f"Карточка создана:\n{output_path}")

    def open_output_folder(self):
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(OUTPUT_DIR)

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


def main():
    root = tk.Tk()
    SizCardApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
