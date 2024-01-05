#все кроме груиипровки по примечаниям
import re
import pandas as pd 
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback


# ==== Utility Functions ====
def read_file(filename):
    try:
        base_name = os.path.basename(filename)  # Например, "ИНРТ.100.01.00.000 Корпус.xlsx"
        
        # Удаляем расширение .xlsx и разделяем имя файла по первому пробелу
        parts = base_name.rsplit('.', 1)[0].split(' ', 1)
        
        # Убедитесь, что parts содержит два элемента
        if len(parts) == 2:
            unit_designation, unit_description = parts
        elif len(parts) == 1:
            unit_designation = parts[0]  # Если нет описания, весь текст является обозначением
        else:
            raise ValueError(f"Непредвиденный формат файла: {filename}")

        # Чтение файла
        df = pd.read_excel(filename, header=None)
        
        return df, unit_designation

    except FileNotFoundError:
        print(f"Ошибка: Файл '{filename}' не найден.")
        return None, None


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

# Загрузка данных материалов из файла
def load_materials_data(materials_file_path):
    try:
        materials_data = pd.read_excel(materials_file_path)
        materials_dict = dict(zip(materials_data['Код'], materials_data['Наименование']))
        return materials_dict
    except Exception as e:
        print(f"Ошибка при загрузке файла материалов: {e}")
        return {}

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
        df, _ = read_file(self.filename + '.xlsx')  # Обратите внимание на использование df, _
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

    def save_tree_to_excel(self, filename="tree_output.xlsx", deal_name="", total_quantity=1):
        lines = []
        self.collect_tree_data(self.assembly_unit, lines)
        # Подготовка данных
        data = []
        for level, name, quantity, note in lines:
            row = [None] * 11
            row[level + 1] = name
            row[-2] = quantity * total_quantity  # Умножаем на количество изделий в сделке
            row[-1] = note
            data.append(row)

        columns = ['Unnamed: 0', 'Наименование', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Структура изделия', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10']
        df = pd.DataFrame(data, columns=columns)
        # Добавляем заголовок с названием главного файла и номером сделки
        df.loc[-1] = [f"Главный файл: {self.assembly_unit.filename}, Сделка: {deal_name}"] + [''] * 10  # Добавление в начало
        df.index = df.index + 1
        df = df.sort_index()  # Сортировка индекса после добавления строки

        df.to_excel(filename, index=False)
        print(f"Дерево сборочных единиц сохранено в {filename}")

    def collect_tree_data(self, unit, lines, level=0):
        lines.append((level, unit.filename, unit.quantity, 'Сборочная единица'))
        for detail in unit.details:
            lines.append((level + 1, detail.designation, detail.total_qty, detail.note))
        for subunit in unit.sub_units:
            self.collect_tree_data(subunit, lines, level + 1)


class DataAggregator:
    def __init__(self, assembly_unit, materials_dict):
        self.assembly_unit = assembly_unit
        self.materials_dict = materials_dict

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
        print(aggregated_data)
        return aggregated_data

    def print_aggregated_data(self, text_widget=None):
        aggregated_data = self.aggregate_details()
        if text_widget:
            for index, row in aggregated_data.iterrows():
                line = f"{row['Обозначение']} - {row['Наименование']}: {row['Общее количество']} {row['Примечание']}\n"
                text_widget.insert(tk.END, line)
        else:
            print(aggregated_data)
    
    def save_aggregated_data(self, aggregated_data, aggregated_file_name):
        try:
            # Попытка сохранения агрегированных данных
            aggregated_data.to_excel(aggregated_file_name, index=False)
            print(f"Итоговые данные сохранены в файл: {aggregated_file_name}")
        except Exception as e:
            print(f"Ошибка при сохранении итоговых данных: {e}")
    

    def save_grouped_data(self, main_filename, directory, deal_name, total_quantity):
        aggregated_data = self.aggregate_details()
        print(f"Сохранение группированных данных в {directory}")

        for note, group in aggregated_data.groupby('Примечание'):
            filename = self.get_filename_from_note(note, main_filename)
            file_path = os.path.join(directory, filename)

            group['Новое количество'] = group['Общее количество'] * total_quantity
            group.insert(0, 'Название сделки', deal_name)

            if os.path.exists(file_path):
                existing_data = pd.read_excel(file_path)
                group = pd.concat([existing_data, group], ignore_index=True)

            if filename.endswith("1с.xlsx"):
                group['Материал'] = group['Примечание'].apply(lambda x: self.match_material(x))

            group.to_excel(file_path, index=False)
            print(f"Файл сохранен: {file_path} с {len(group)} строками")



  
    
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
            "М": "Механообработка",
            "Л": "Лазер",
            "Г": "Гибка",
            "К": "Покраска порошковая",
            "КЭ": "Покраска эмаль",
            "Ц": "Цинкование",
            "Р": "Гидроабразив резина",
            "В": "Вальцовка",
            "ГО": "Гидроабразив металл",
            "П": "Литье полиуретана",
            "И": "ИНЕРТА",
            "ИЗ": "ИЗПА",
            "3d": "3d-печать"
                # Убедитесь, что ключи и значения уникальны и соответствуют вашим данным
            }

    def match_material(self, note):
        nf_codes = re.findall(r'НФ-\d{8}', note)
        materials = [self.materials_dict.get(code) for code in nf_codes]
        return ', '.join(filter(None, materials))

    def get_filename_from_note(self, note, main_filename):
        # Проверка наличия НФ-кодов
        if any("НФ-" in part for part in note.split(',')):
            return f"{main_filename} – 1с.xlsx"

        # Обработка остальных примечаний
        special_notes = []
        for part in note.split(','):
            part = part.strip()
            if part in self.note_to_filename:  # Используйте словарь для получения названия файла из примечания
                special_notes.append(self.note_to_filename[part])
        if special_notes:
            return f"{main_filename} – {' & '.join(special_notes)}.xlsx"

        return f"{main_filename} – Другое.xlsx"  # Стандартное имя, если нет совпадений
   



# ==== GUI Application ====
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Обработчик сборочных единиц")
        self.geometry("800x600")
        # Инициализация materials_dict здесь
        self.materials_dict = None  # Заполняется позже
        self.create_widgets()

    def create_widgets(self):
        # Button for choosing the main assembly file
        self.choose_button = tk.Button(self, text="Выбрать файл", command=self.choose_file)
        self.choose_button.pack(pady=20)

        # Button for choosing the Materials 1С file
        self.choose_materials_button = tk.Button(self, text="Выбрать файл материалов 1С", command=self.choose_materials_file)
        self.choose_materials_button.pack(pady=20)

        # Label and input field for deal number
        self.deal_label = tk.Label(self, text="Номер сделки:")
        self.deal_label.pack()
        self.deal_entry = tk.Entry(self)
        self.deal_entry.pack()

        # Label and input field for the number of items in the deal
        self.quantity_label = tk.Label(self, text="Количество изделий в сделке:")
        self.quantity_label.pack()
        self.quantity_entry = tk.Entry(self)
        self.quantity_entry.pack()

        # Text widget for displaying data
        self.text = tk.Text(self, height=30, width=90)
        self.text.pack(pady=20)

    def choose_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                filename_without_extension = os.path.splitext(os.path.basename(file_path))[0]
                directory = filename_without_extension  # The folder is named after the main file

                if not os.path.exists(directory):
                    os.makedirs(directory)  # Create the folder if it does not exist

                # Retrieve and validate total_quantity
                total_quantity_str = self.quantity_entry.get()
                if total_quantity_str.isdigit() and total_quantity_str != "":
                    total_quantity = int(total_quantity_str)
                else:
                    raise ValueError("The entered quantity is not a valid number or is empty.")

                # Retrieve deal name
                deal_name = self.deal_entry.get()  # Assuming it's okay to be any string

                # Process the file
                main_assembly = AssemblyUnit(filename_without_extension)
                main_assembly.process_file()

                # Print tree structure in the text widget
                data_output = DataOutput(main_assembly)
                self.text.insert(tk.END, "Иерархия сборочных единиц и деталей:\n")
                data_output.print_tree(target=self.text)

                # Aggregate and print details data
                aggregator = DataAggregator(main_assembly, self.materials_dict)
                aggregated_data = aggregator.aggregate_details()
                self.text.insert(tk.END, "\nИтоговые данные:\n")
                aggregator.print_aggregated_data(text_widget=self.text)
               # Save aggregated data
                aggregated_file_name = os.path.join(directory, f"{filename_without_extension}_aggregated.xlsx")
                aggregator.save_aggregated_data(aggregated_data, aggregated_file_name)

                # Process and save grouped data based on notes
                aggregator.save_grouped_data(filename_without_extension, directory, deal_name, total_quantity)

                # Save tree structure to Excel
                tree_file_name = os.path.join(directory, f"{filename_without_extension}_tree_structure.xlsx")
                data_output.save_tree_to_excel(tree_file_name, deal_name, total_quantity)

                messagebox.showinfo("Success", "Files processed and saved successfully!")
            except ValueError as ve:
                messagebox.showerror("Input Error", str(ve))
            except Exception as e:
                traceback.print_exc()  # Prints full traceback of the error
                messagebox.showerror("Error", str(e))



    # Inside the Application class's choose_materials_file method
    def choose_materials_file(self):
        materials_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if materials_file_path:
            try:
                # Загрузка и сохранение данных из файла материалов 1С
                self.materials_dict = load_materials_data(materials_file_path)

                # Выводим сообщение о том, что файл материалов был выбран и загружен
                messagebox.showinfo("File Chosen", f"Materials file chosen: {materials_file_path}")
            except Exception as e:
                traceback.print_exc()  # Печать полной трассировки ошибки
                messagebox.showerror("Error", str(e))



if __name__ == "__main__":
    app = Application()
    app.mainloop()

