#классы модули математика дерево gui короче все вместе 

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# === Модуль обработки данных ===
class AssemblyUnit:
    def __init__(self, filename, quantity=1, level=0, parent_qty=1):
        self.filename = filename
        self.quantity = quantity
        self.level = level
        self.parent_qty = parent_qty
        self.details = []
        self.sub_units = []

    def process_file(self):
        df = read_file(self.filename + '.xlsx')
        if df is not None:
            self.process_details(df)
            self.process_sub_units(df)

    def process_details(self, df):
        details_section = extract_section(df, 'Детали')
        for _, row in details_section.iterrows():
            total_qty = row[5] * self.parent_qty
            detail = Detail(row[3], row[4], row[5], row[6], total_qty)
            self.details.append(detail)

    def process_sub_units(self, df):
        sub_units_section = extract_section(df, 'Сборочные единицы')
        for _, row in sub_units_section.iterrows():
            sub_unit = AssemblyUnit(row[3], row[5], self.level + 1, self.quantity * self.parent_qty)
            sub_unit.process_file()
            self.sub_units.append(sub_unit)

class Detail:
    def __init__(self, designation, name, quantity, note, total_qty):
        self.designation = designation
        self.name = name
        self.quantity = quantity
        self.note = note
        self.total_qty = total_qty


# Utility functions
def list_xlsx_files():
    files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    return files

def read_file(filename):
    try:
        df = pd.read_excel(filename, header=None)
        return df
    except FileNotFoundError:
        print(f"Ошибка: Файл '{filename}' не найден.")
        return None

def extract_section(df, start_label, end_label=None):
    try:
        start_idx = df[df.iloc[:, 4] == start_label].index[0] + 1
        if end_label:
            end_idx = df.index[df.iloc[:, 4] == end_label][0]
            return df.iloc[start_idx:end_idx].dropna(subset=[5])
        return df.iloc[start_idx:].dropna(subset=[5])
    except IndexError:
        print(f"Секция '{start_label}' не найдена.")
        return pd.DataFrame()

# === Модуль вывода данных ===
class DataOutput:
    def __init__(self, assembly_unit):
        self.assembly_unit = assembly_unit

    def print_tree(self, unit=None, level=0, target=None):
        if unit is None:
            unit = self.assembly_unit
        indent = "  " * level
        tree_text = f"{indent}{unit.filename} x{unit.quantity} (Сборочная единица на уровне {level}):\n"
        if target:
            target.insert(tk.END, tree_text)

        for detail in unit.details:
            detail_text = f"{indent}  - {detail.designation}: {detail.name}, Количество: {detail.total_qty}, Примечание: {detail.note}\n"
            if target:
                target.insert(tk.END, detail_text)

        for subunit in unit.sub_units:
            self.print_tree(subunit, level + 1, target)

# === Модуль сортировки и агрегации данных ===
class DataAggregator:
    def __init__(self, assembly_unit):
        self.assembly_unit = assembly_unit

    def get_details(self, unit=None):
        if unit is None:
            unit = self.assembly_unit
        details = []
        for detail in unit.details:
            details.append(detail)
        for subunit in unit.sub_units:
            details.extend(self.get_details(subunit))
        return details

    def aggregate_details(self):
        details = self.get_details()
        df = pd.DataFrame([{
            "Обозначение": detail.designation,
            "Наименование": detail.name,
            "Количество": detail.total_qty,
            "Примечание": detail.note
        } for detail in details])
        
        # Группировка и суммирование количества
        aggregated_data = df.groupby(["Обозначение", "Наименование", "Примечание"])['Количество'].sum().reset_index()
        return aggregated_data

    def print_aggregated_data(self, text_widget=None):
        aggregated_data = self.aggregate_details()
        if text_widget:
            for index, row in aggregated_data.iterrows():
                line = f"{row['Обозначение']} - {row['Наименование']}: {row['Количество']} {row['Примечание']}\n"
                text_widget.insert(tk.END, line)

    def save_grouped_data(self, main_filename):
        aggregated_data = self.aggregate_details()
        for note, group in aggregated_data.groupby('Примечание'):
            filename = self.get_filename_from_note(note, main_filename)
            group.to_excel(filename, index=False)
            print(f"Сохранено {len(group)} строк в файл: {filename}")
    
    def get_filename_from_note(self, note, main_filename):
        # Обновленный словарь для соответствия примечаний и расшифровок
        note_to_filename = {
            "Р": "Резка",
            "С": "Сварка",
            "ЛГ": "Лазер",
            "В": "Вальцовка",
            "ГР": "Гидроабразив резины",
            "ГМ": "Гидроабразив металла",
            "Ф": "Фрезеровка",
            "Ц": "Токарка",
            "3Д": "3D-печать",
            "КЭ": "ЛКП эмаль",
            "КП": "ЛКП порошок",
            "Б": "Балансировка",
            "ТО": "Термообработка",
            "ФМ": "Формовка",
            "Ш": "Шлифовка",
            "П": "Пассивация",
            # ... добавьте другие примечания по мере необходимости ...
        }

        # Логика обработки примечаний
        base_filename = main_filename.split('.')[0]
        special_notes = []  # Для хранения специальных примечаний, например, материала

        # Обработка стандартных и специальных примечаний
        for part in note.split(','):
            if part.startswith("НФ-"):  # Обрабатываем специальное примечание материала
                special_notes.append(part)
            else:  # Обрабатываем стандартные примечания
                if part in note_to_filename:
                    return f"{base_filename} – {note_to_filename[part]}.xlsx"
        
        # Если были специальные примечания, обрабатываем их
        if special_notes:
            return f"{base_filename} – {', '.join(special_notes)}.xlsx"

        return f"{base_filename} – Другое.xlsx"  # Если ни одно примечание не совпало



# === GUI using Tkinter ===
class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Обработчик сборочных единиц")
        self.geometry("800x600")
        self.create_widgets()

    def create_widgets(self):
        self.choose_button = tk.Button(self, text="Выбрать файл", command=self.choose_file)
        self.choose_button.pack(pady=20)

        self.text = tk.Text(self, height=30, width=90)
        self.text.pack(pady=20)

    def choose_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                filename_without_extension = os.path.splitext(os.path.basename(file_path))[0]
                main_assembly = AssemblyUnit(filename_without_extension)
                main_assembly.process_file()

                # Вывод иерархии
                data_output = DataOutput(main_assembly)
                self.text.insert(tk.END, "Иерархия сборочных единиц и деталей:\n")
                data_output.print_tree(target=self.text)

                # Агрегация и вывод итоговых данных
                aggregator = DataAggregator(main_assembly)
                self.text.insert(tk.END, "\nИтоговые данные:\n")
                aggregator.print_aggregated_data()
                aggregator.save_grouped_data(filename_without_extension)

                messagebox.showinfo("Обработка завершена", "Данные успешно обработаны!")

            except Exception as e:
                messagebox.showerror("Ошибка", str(e))

def main():
    app = Application()
    app.mainloop()

if __name__ == "__main__":
    main()

