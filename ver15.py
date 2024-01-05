# -*- coding: utf-8 -*-

import re
import pandas as pd 
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback
import glob

def log_message(message):
    """Print a log message to the console."""
    print(f"LOG: {message}")

# ==== Utility Functions ====
import os
import glob


import os

import os

def find_file(designation, parent_directory):
    parent_directory = os.path.normpath(parent_directory)
    
    # Преобразование обозначения в строку, на случай, если оно было интерпретировано как число
    designation_str = str(designation)
    
    exact_path = os.path.join(parent_directory, designation_str + '.xlsx')
    if os.path.exists(exact_path):
        return exact_path
    
    for file in os.listdir(parent_directory):
        if file.startswith(designation_str) and file.endswith('.xlsx'):
            return os.path.join(parent_directory, file)
    
    return None







import os
import pandas as pd

import os
import pandas as pd

def read_file(filename, parent_directory=None):
    try:
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        full_path = os.path.join(parent_directory, filename) if parent_directory else filename
        log_message(f"Full path to read: {full_path}")

        if not os.path.exists(full_path):
            log_message(f"Error: File not found at path: {full_path}")
            return None, None
        if not os.path.isfile(full_path):
            log_message(f"Error: Expected a file but got a directory or non-existent path: {full_path}")
            return None, None

        df = pd.read_excel(full_path, engine='openpyxl')  # Использование 'openpyxl' для кириллических имен
        unit_designation = os.path.splitext(os.path.basename(full_path))[0]

        log_message(f"File read successfully: {full_path}")
        return df, unit_designation

    except Exception as e:
        log_message(f"Unexpected error when reading file: {e}")
        return None, None











   



def extract_section(df, start_label, end_label=None):
    try:
        # Находим индекс строки, которая начинается с start_label
        start_idx = df.index[df.iloc[:, 4] == start_label].tolist()
        if not start_idx:  # Если такой строки нет, возвращаем пустой DataFrame
            log_message(f"Section '{start_label}' not found.")
            return pd.DataFrame()
        else:
            start_idx = start_idx[0] + 1  # Сдвигаем на одну строку вниз

        if end_label:  # Если указан end_label, находим его индекс
            end_idx = df.index[df.iloc[:, 4] == end_label].tolist()
            if not end_idx:
                log_message(f"End section '{end_label}' not found. Extracting till the end.")
                return df.iloc[start_idx:].dropna(how='all')  # Возвращаем все до конца, если не найден end_label
            else:
                end_idx = end_idx[0]
                return df.iloc[start_idx:end_idx].dropna(how='all')
        else:  # Если end_label не указан, возвращаем все до конца файла
            return df.iloc[start_idx:].dropna(how='all')

    except Exception as e:
        log_message(f"Error extracting section: {e}")
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

    def process_file(self, parent_directory=None):
        file_path = find_file(self.filename, parent_directory)  # Используйте find_file для получения полного пути
        if file_path:
            df, unit_designation = read_file(file_path, parent_directory)  # Передайте полный путь
            if df is not None:
                self.process_details(df)
                self.process_sub_units(df, parent_directory)
        else:
            log_message(f"Error: File '{self.filename}' not found in directory '{parent_directory}'")


    def process_sub_units(self, df, parent_directory):
        sub_units_section = extract_section(df, 'Сборочные единицы')
        for _, row in sub_units_section.iterrows():
            designation = str(row.iloc[3])
            file_name = find_file(designation, parent_directory)
            if file_name:
                sub_unit = AssemblyUnit(file_name, float(row.iloc[5]), self.level + 1, self.quantity * self.parent_qty)
                sub_unit.process_file(parent_directory)
                self.sub_units.append(sub_unit)
            else:
                log_message(f"Error: File for sub-unit '{designation}' not found.")



    def process_details(self, df):
        details_section = extract_section(df, 'Детали')
        for _, row in details_section.iterrows():
            total_qty = float(row.iloc[5]) * self.parent_qty * self.quantity
            detail = Detail(
                str(row.iloc[3]),
                str(row.iloc[4]),
                row.iloc[5],
                str(row.iloc[6]),
                total_qty
            )
            self.details.append(detail)






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
    

    def save_1c_data(self, data_1c, file_path_1c):
        log_message(f"Starting to save 1c data to {file_path_1c}")
        try:
            print("Full path to the 1c.xlsx file:", os.path.abspath(file_path_1c))
            if os.path.exists(file_path_1c):
                existing_data = pd.read_excel(file_path_1c)
                log_message("Existing data found, updating...")
                updated_data = pd.concat([existing_data, data_1c]).drop_duplicates(keep='first')
                updated_data.to_excel(file_path_1c, index=False)
            else:
                log_message("No existing file found, creating a new one...")
                data_1c.to_excel(file_path_1c, index=False)
            log_message(f"Data successfully saved to {file_path_1c}")
        except Exception as e:
            log_message(f"Error saving data to 1C: {e}")





    def save_grouped_data(self, main_filename, directory, deal_name, total_quantity):
        # Получение агрегированных данных
        aggregated_data = self.aggregate_details()

        # Словарь для группированных данных по типу файла
        grouped_data = {}

        # Проход по всем строкам агрегированных данных
        for _, row in aggregated_data.iterrows():
            # Получение имени файла на основе примечания
            note = row['Примечание']
            filename = self.get_filename_from_note(note, main_filename)
            file_path = os.path.join(directory, filename)  # Исправлено: использование directory для пути

            # Добавление данных в соответствующую группу
            if file_path not in grouped_data:
                grouped_data[file_path] = []
            grouped_data[file_path].append(row)

        # Сохранение данных в соответствующие файлы
        for file_path, data_rows in grouped_data.items():
            # Преобразование списка данных обратно в DataFrame
            df_to_save = pd.DataFrame(data_rows)

            print("Текущий путь к файлу для сохранения:", file_path)
            if "1с.xlsx" in file_path:
                print("Сохраняем данные в файл 1c:", file_path)
                print("Данные для сохранения:", df_to_save)
                # Убедитесь, что убраны неиспользуемые строки, связанные с existing_data
                if os.path.exists(file_path):
                    existing_data = pd.read_excel(file_path)
                    updated_data = pd.concat([existing_data, df_to_save]).drop_duplicates(keep='first')
                    updated_data.to_excel(file_path, index=False)
                else:
                    df_to_save.to_excel(file_path, index=False)
            else:
                # Обычное сохранение данных для других типов файлов
                df_to_save.to_excel(file_path, index=False)
                print(f"Файл сохранен: {file_path} с {len(df_to_save)} строками.")









  
    
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
}   


    def match_material(self, note):
        nf_codes = re.findall(r'НФ-\d{8}', note)
        materials = [self.materials_dict.get(code) for code in nf_codes]
        return ', '.join(filter(None, materials))

    def get_filename_from_note(self, note, main_filename):
        filename = f"{main_filename} – Другое.xlsx"  # Стандартное имя, если нет совпадений

        for part in note.split(','):
            part = part.strip()  # Удаление лишних пробелов
            key = part.rstrip('0123456789')  # Извлечение ключа (без цифр очереди)

            if key in self.note_to_filename:  # Проверка наличия ключа в словаре
                name = self.note_to_filename[key]  # Получение имени процесса

                # Проверка наличия и обработка номера очереди
                if len(part) > len(key):  # Если есть цифры после ключа
                    queue_num = part[len(key):]  # Извлечение номера очереди
                    filename = f"{main_filename} – {name} {queue_num}-я очередь.xlsx"
                else:  # Если нет номера очереди, то просто имя процесса
                    filename = f"{main_filename} – {name}.xlsx"

        return filename



# ==== GUI Application ====
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Обработчик сборочных единиц")
        self.geometry("800x600")
        self.materials_dict = None  # Заполняется позже
        self.create_widgets()

    def create_widgets(self):
        # Button for choosing the Materials 1С file
        self.choose_materials_button = tk.Button(self, text="Выбрать файл материалов 1С", command=self.choose_materials_file)
        self.choose_materials_button.pack(pady=20)

        # Button for choosing the main assembly file
        self.choose_button = tk.Button(self, text="Выбрать главный файл", command=self.choose_file)
        self.choose_button.pack(pady=20)

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
        log_message("Choosing file via dialog")
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            log_message(f"File chosen: {file_path}")
            try:
                filename_without_extension = os.path.splitext(os.path.basename(file_path))[0]
                directory = os.path.dirname(file_path)  # Старая директория
                output_directory = os.path.join(directory, "Output")  # Новая поддиректория для вывода
                if not os.path.exists(output_directory):
                    os.makedirs(output_directory)

                total_quantity_str = self.quantity_entry.get()
                if total_quantity_str.isdigit() and total_quantity_str != "":
                    total_quantity = int(total_quantity_str)
                else:
                    raise ValueError("The entered quantity is not a valid number or is empty.")

                deal_name = self.deal_entry.get()

                main_assembly = AssemblyUnit(filename_without_extension)
                main_assembly.process_file(directory)  # Используйте directory для чтения исходного файла

                data_output = DataOutput(main_assembly)
                self.text.insert(tk.END, "Hierarchy of assembly units and details:\n")
                data_output.print_tree(target=self.text)

                aggregator = DataAggregator(main_assembly, self.materials_dict)
                aggregated_data = aggregator.aggregate_details()
                self.text.insert(tk.END, "\nFinal data:\n")
                aggregator.print_aggregated_data(text_widget=self.text)

                aggregated_file_name = os.path.join(output_directory, f"{filename_without_extension}_aggregated.xlsx")
                aggregator.save_aggregated_data(aggregated_data, aggregated_file_name)

                aggregator.save_grouped_data(filename_without_extension, output_directory, deal_name, total_quantity)

                tree_file_name = os.path.join(output_directory, f"{filename_without_extension}_tree_structure.xlsx")
                data_output.save_tree_to_excel(tree_file_name, deal_name, total_quantity)

                messagebox.showinfo("Success", "Files processed and saved successfully!")
                log_message("File processing completed successfully")

            except ValueError as ve:
                log_message(f"Input Error: {ve}")
                messagebox.showerror("Input Error", str(ve))
            except Exception as e:
                log_message(f"Error processing file: {e}")
                traceback.print_exc()  # Prints full traceback of the error
                messagebox.showerror("Error", str(e))






    # Inside the Application class's choose_materials_file method
    def choose_materials_file(self):
        materials_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if materials_file_path:
            if os.path.exists(materials_file_path):
                try:
                    self.materials_dict = load_materials_data(materials_file_path)
                    messagebox.showinfo("File Chosen", f"Materials file chosen: {materials_file_path}")
                    print("Materials Dictionary:", self.materials_dict)
                except Exception as e:
                    traceback.print_exc()  # Печать полной трассировки ошибки
                    messagebox.showerror("Error", str(e))
            else:
                messagebox.showerror("Error", "Файл материалов 1С не найден.")




if __name__ == "__main__":
    app = Application()
    app.mainloop()


