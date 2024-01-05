#gui верн итог данные + дерево
import pandas as pd 
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# ==== Utility Functions ====
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
            end_idx = df[df.iloc[:, 4] == end_label].index[0]
            return df.iloc[start_idx:end_idx].dropna(subset=[5])
        return df.iloc[start_idx:].dropna(subset=[5])
    except IndexError:
        print(f"Секция '{start_label}' не найдена.")
        return pd.DataFrame()

# ==== Core Classes ====
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
            total_qty = row[5] * self.parent_qty * self.quantity  # Умножаем на количество родителя и текущее количество
            detail = Detail(row[3], row[4], row[5], row[6], total_qty)
            self.details.append(detail)


    def process_sub_units(self, df):
        sub_units_section = extract_section(df, 'Сборочные единицы')
        for _, row in sub_units_section.iterrows():
            # Аналогично, используйте числовые индексы для доступа к данным
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


class DataOutput:
    def __init__(self, assembly_unit):
        self.assembly_unit = assembly_unit

    def print_tree(self, unit=None, level=0, target=None):
        if unit is None:
            unit = self.assembly_unit
        indent = "  " * level
        tree_text = f"{indent}{unit.filename} x{unit.quantity} (Уровень {level}):\n"
        if target:
            target.insert(tk.END, tree_text)
        else:
            print(tree_text)

        for detail in unit.details:
            detail_text = f"{indent}  - {detail.designation}: {detail.name}, Кол-во: {detail.total_qty}, Примечание: {detail.note}\n"
            if target:
                target.insert(tk.END, detail_text)
            else:
                print(detail_text)

        for subunit in unit.sub_units:
            self.print_tree(subunit, level + 1, target)

    def save_tree_to_excel(self, filename="tree_output.xlsx"):
        lines = []
        self.collect_tree_data(self.assembly_unit, lines)
        df = pd.DataFrame(lines, columns=['Уровень', 'Сборочная единица', 'Количество', 'Примечание'])
        df.to_excel(filename, index=False)
        print(f"Дерево сборочных единиц сохранено в {filename}")

    def collect_tree_data(self, unit, lines, level=0):
        lines.append([level, unit.filename, unit.quantity, 'Сборочная единица'])
        for detail in unit.details:
            lines.append([level + 1, detail.designation, detail.total_qty, detail.note])
        for subunit in unit.sub_units:
            self.collect_tree_data(subunit, lines, level + 1)


class DataAggregator:
    def __init__(self, assembly_unit):
        self.assembly_unit = assembly_unit

    def get_details(self, unit=None):
        if unit is None:
            unit = self.assembly_unit
        details = unit.details.copy()
        for subunit in unit.sub_units:
            details.extend(self.get_details(subunit))
        return details

    def aggregate_details(self):
        all_details = self.get_details()
        details_data = []
        for detail in all_details:
            details_data.append([detail.designation, detail.name, detail.quantity, detail.total_qty, detail.note])

        # Создаем DataFrame и указываем точные названия столбцов
        df = pd.DataFrame(details_data, columns=['Обозначение', 'Наименование', 'Количество на единицу', 'Общее количество', 'Примечание'])
        # Группируем данные по 'Обозначение', 'Наименование', и 'Примечание', суммируя 'Общее количество'
        aggregated_data = df.groupby(['Обозначение', 'Наименование', 'Примечание'])['Общее количество'].sum().reset_index()
        return aggregated_data

    def print_aggregated_data(self, text_widget=None):
        aggregated_data = self.aggregate_details()
        if text_widget:
            for index, row in aggregated_data.iterrows():
                line = f"{row['Обозначение']} - {row['Наименование']}: {row['Общее количество']} {row['Примечание']}\n"
                text_widget.insert(tk.END, line)
        else:
            print(aggregated_data)

    def save_grouped_data(self, main_filename):
        aggregated_data = self.aggregate_details()
        aggregated_data.to_excel(f"{main_filename}_aggregated.xlsx", index=False)
        print(f"Сохранено в {main_filename}_aggregated.xlsx")


# ==== GUI Application ====
import tkinter as tk
from tkinter import filedialog, messagebox

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Обработчик сборочных единиц")
        self.geometry("800x600")
        self.create_widgets()

    def create_widgets(self):
        # Кнопка для выбора файла
        self.choose_button = tk.Button(self, text="Выбрать файл", command=self.choose_file)
        self.choose_button.pack(pady=20)

        # Текстовый виджет для отображения данных
        self.text = tk.Text(self, height=30, width=90)
        self.text.pack(pady=20)

    def choose_file(self):
        # Диалоговое окно для выбора файла
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
                aggregator.print_aggregated_data(text_widget=self.text)
                aggregator.save_grouped_data(filename_without_extension)

                # Всплывающее окно с информацией об успешной обработке данных
                messagebox.showinfo("Обработка завершена", "Данные успешно обработаны!")
            except Exception as e:
                # Всплывающее окно с информацией об ошибке
                messagebox.showerror("Ошибка", str(e))

def main():
    app = Application()
    app.mainloop()

if __name__ == "__main__":
    main()

