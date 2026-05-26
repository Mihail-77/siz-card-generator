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
    "Табельный номер",
    "ФИО",
    "Дата приема",
    "Пол",
    "Рост",
    "Размер одежды",
    "Размер обуви",
    "Размер головного убора",
    "Размер перчаток",
]


class SizCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор карточек СИЗ")
        self.root.resizable(False, False)

        self.department_var = tk.StringVar()
        self.position_var = tk.StringVar()
        self.field_vars = {label: tk.StringVar() for label in FIELD_LABELS}

        self.build_form()
        self.load_departments()

    def build_form(self):
        frame = ttk.Frame(self.root, padding=16)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frame, text="Подразделение").grid(row=0, column=0, sticky="w", pady=4)
        self.department_combo = ttk.Combobox(
            frame,
            textvariable=self.department_var,
            state="readonly",
            width=42,
        )
        self.department_combo.grid(row=0, column=1, sticky="ew", pady=4)
        self.department_combo.bind("<<ComboboxSelected>>", self.on_department_selected)

        ttk.Label(frame, text="Должность").grid(row=1, column=0, sticky="w", pady=4)
        self.position_combo = ttk.Combobox(
            frame,
            textvariable=self.position_var,
            state="readonly",
            width=42,
        )
        self.position_combo.grid(row=1, column=1, sticky="ew", pady=4)

        for index, label in enumerate(FIELD_LABELS, start=2):
            ttk.Label(frame, text=label).grid(row=index, column=0, sticky="w", pady=4)
            entry = ttk.Entry(frame, textvariable=self.field_vars[label], width=45)
            entry.grid(row=index, column=1, sticky="ew", pady=4)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=len(FIELD_LABELS) + 2, column=0, columnspan=2, sticky="e", pady=(12, 0))

        ttk.Button(button_frame, text="Открыть папку output", command=self.open_output_folder).grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(button_frame, text="Создать карточку", command=self.create_card).grid(row=0, column=1)

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

        self.department_combo["values"] = departments
        if departments:
            self.department_var.set(departments[0])
            self.update_positions(departments[0])
        else:
            messagebox.showerror(
                "Нет подразделений",
                "В data/norms.xlsx не найдены подразделения. Проверьте файл норм.",
            )

    def on_department_selected(self, _event=None):
        self.update_positions(self.department_var.get())

    def update_positions(self, department):
        try:
            positions = get_positions_by_department(NORMS_PATH, department)
        except Exception as error:
            messagebox.showerror("Ошибка", f"Не удалось прочитать должности:\n{error}")
            positions = []

        self.position_combo["values"] = positions
        self.position_var.set(positions[0] if positions else "")
        if department and not positions:
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
            employee[label] = variable.get().strip()

        return employee

    def validate_employee(self, employee):
        missing = [field for field, value in employee.items() if not value]
        if missing:
            messagebox.showerror("Заполните поля", "Не заполнены поля:\n" + "\n".join(missing))
            return False

        if not self.position_combo["values"]:
            messagebox.showerror(
                "Не найдены нормы",
                "Для выбранного подразделения нет должностей в norms.xlsx.",
            )
            return False

        return True

    def create_card(self):
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
            messagebox.showerror("Ошибка", f"Не удалось создать карточку:\n{error}")
            return

        if output_path is None:
            messagebox.showerror(
                "Не найдены нормы",
                "Для выбранного подразделения и должности нормы не найдены.",
            )
            return

        messagebox.showinfo("Готово", f"Карточка создана:\n{output_path}")

    def open_output_folder(self):
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(OUTPUT_DIR)


def main():
    root = tk.Tk()
    SizCardApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
