#работет 
import re
import pandas as pd 
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback
import argparse
import glob
import datetime


def log_message(message):
    """Print a log message to the console."""
    print(f"LOG: {message}")

def read_file(filename):
    log_message(f"Attempting to read file: {filename}")
    try:
        df = pd.read_excel(filename, header=None)
        log_message(f"File read successfully: {filename}")

        # Получение обозначения из названия файла
        base_name = os.path.splitext(os.path.basename(filename))[0]
        unit_designation = base_name.split(' ')[0]

        return df, unit_designation
    except FileNotFoundError:
        log_message(f"FileNotFoundError: File '{filename}' not found.")
    except pd.errors.EmptyDataError:
        log_message(f"EmptyDataError: File '{filename}' is empty or invalid.")
    except Exception as e:
        log_message(f"Unexpected error when reading file '{filename}': {e}")

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


def process_assembly_unit(main_file, quantity, materials_file, deal_name, output_directory):
    log_message(f"Starting processing assembly unit for main file: {main_file}")
    main_file_path = os.path.abspath(main_file)
    materials_file_path = os.path.abspath(materials_file)

    # Изменение формата даты и времени
    current_datetime = datetime.datetime.now().strftime("%H-%M_%d-%m-%y")
    output_directory = os.path.join(output_directory, current_datetime + "_" + os.path.basename(main_file_path))

    log_message(f"Output directory: {output_directory}")
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
        log_message(f"Output directory created: {output_directory}")

    log_message("Loading materials data")
    materials_dict = load_materials_data(materials_file_path)

    main_assembly = AssemblyUnit(os.path.splitext(os.path.basename(main_file_path))[0], quantity)
    main_assembly.process_file(materials_dict)

    data_output = DataOutput(main_assembly)
    tree_output_path = os.path.join(output_directory, f"{main_assembly.filename}_tree_structure.xlsx")
    data_output.save_tree_to_excel(tree_output_path, deal_name, quantity)

    aggregator = DataAggregator(main_assembly, materials_dict)
    aggregated_data = aggregator.aggregate_details()
    aggregated_file_name = os.path.join(output_directory, f"{main_assembly.filename}_aggregated.xlsx")
    aggregator.save_aggregated_data(aggregated_data, aggregated_file_name, deal_name, quantity)

    file_path_1c = os.path.join(output_directory, f"{main_assembly.filename}_1C.xlsx")
    aggregator.save_1c_data(aggregated_data, file_path_1c, quantity)

    aggregator.save_grouped_data(main_file_path, output_directory, deal_name, quantity)

    log_message(f"Processing completed, results saved in {output_directory}")









# ==== Core Classes ====
class AssemblyUnit:

    def __init__(self, filename, quantity=1, level=0, parent_qty=1, designation='', name='', note=''):
        self.filename = filename
        self.quantity = quantity
        self.level = level
        self.parent_qty = parent_qty
        self.details = []
        self.sub_units = []
        self.designation = designation
        self.name = name
        self.note = note  # Добавлен атрибут note

    def process_file(self, materials_dict):
        log_message(f"Processing file for unit: {self.filename}")
        
        # Проверяем существование файла с полным названием
        full_filename = f"{self.filename}.xlsx"
        df, unit_designation = read_file(full_filename)

        if df is None:
            log_message(f"File not found with name: {self.filename}. Trying glob search.")
            possible_files = glob.glob(self.filename + '*')
            for file in possible_files:
                if os.path.isfile(file):
                    log_message(f"Trying file found by glob: {file}")
                    df, unit_designation = read_file(file)
                    if df is not None:
                        log_message(f"File found by glob: {file}")
                        break

        if df is not None:
            self.designation = unit_designation
            log_message(f"Processing details and sub-units for file: {self.filename}")
            self.process_details(df, materials_dict)
            self.process_sub_units(df, materials_dict)
        else:
            log_message(f"Error: Unable to find the file for '{self.filename}'.")



 





    def process_details(self, df, materials_dict):
        details_section = extract_section(df, 'Детали')
        for _, row in details_section.iterrows():
            total_qty = row[5] * self.parent_qty * self.quantity
            detail = Detail(row[3], row[4], row[5], row[6], total_qty, materials_dict)
            self.details.append(detail)

    def process_sub_units(self, df, materials_dict):
        sub_units_section = extract_section(df, 'Сборочные единицы')
        for _, row in sub_units_section.iterrows():
            sub_unit = AssemblyUnit(row[3], row[5], self.level + 1, self.quantity * self.parent_qty)
            sub_unit.process_file(materials_dict)  # передаем materials_dict
            self.sub_units.append(sub_unit)







class Detail:
    def __init__(self, designation, name, quantity, note, total_qty, materials_dict):
        self.designation = designation
        self.name = name
        self.quantity = quantity
        self.note = note
        self.total_qty = total_qty
        self.materials = self.match_material(note, materials_dict)

    def match_material(self, note, materials_dict):
        # Проверяем, что примечание является строкой, иначе преобразуем или заменяем
        if not isinstance(note, str):
            if pd.isna(note):  # Если note равно NaN
                note = ''  # Заменяем на пустую строку
            else:
                note = str(note)  # Преобразуем в строку

        # Поиск кодов материалов и их соответствия
        nf_codes = re.findall(r'НФ-\d{8}', note)
        materials = [materials_dict.get(code, "Неизвестный материал") for code in nf_codes]
        return ', '.join(filter(None, materials))








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

        data = []
        for level, designation, name, quantity, note in lines:
            row = [''] * 7 + [None] * 6  # Адаптируйте количество столбцов под вашу структуру
            row[level] = '--' * level + designation  # Иерархия и обозначение

            row[7] = level  # Уровень
            row[8] = name  # Наименование
            row[9] = quantity  # Количество
            row[10] = quantity * total_quantity  # Общее количество
            row[11] = deal_name  # Название сделки
            row[12] = note  # Примечание

            data.append(row)

        columns = ['Иерархия 1', 'Иерархия 2', 'Иерархия 3', 'Иерархия 4', 'Иерархия 5', 'Иерархия 6', 'Иерархия 7', 'Уровень', 'Наименование', 'Количество', 'Общее количество', 'Название сделки', 'Примечание']
        df = pd.DataFrame(data, columns=columns)

        print(f"Сохранение дерева сборочных единиц в {filename}")
        df.to_excel(filename, index=False)
        print(f"Дерево сборочных единиц успешно сохранено в {filename}")



    def collect_tree_data(self, unit, lines, level=0):
        # Если объект является сборочной единицей, обрабатываем его детали и подсборки
        if isinstance(unit, AssemblyUnit):
            lines.append((level, unit.designation, unit.name, unit.quantity, 'Сборочная единица'))

            for detail in unit.details:
                self.collect_tree_data(detail, lines, level + 1)

            for subunit in unit.sub_units:
                self.collect_tree_data(subunit, lines, level + 1)
        
        # Если объект является деталью, добавляем информацию о детали
        elif isinstance(unit, Detail):
            lines.append((level, unit.designation, unit.name, unit.total_qty, unit.note))






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
            details_data.append([detail.designation, detail.name, detail.note, detail.materials, detail.total_qty])
        df = pd.DataFrame(details_data, columns=['Обозначение', 'Наименование', 'Примечание', 'Материалы', 'Общее количество'])
        aggregated_data = df.groupby(['Обозначение', 'Наименование', 'Примечание', 'Материалы'])['Общее количество'].sum().reset_index()
        return aggregated_data


    def print_aggregated_data(self, text_widget=None):
        aggregated_data = self.aggregate_details()
        if text_widget:
            for index, row in aggregated_data.iterrows():
                line = f"{row['Обозначение']} - {row['Наименование']}: {row['Общее количество']} {row['Примечание']}\n"
                text_widget.insert(tk.END, line)
        else:
            print(aggregated_data)



    
    def save_aggregated_data(self, aggregated_data, aggregated_file_name, deal_name, total_quantity):
        try:
            # Добавление информации о сделке и расчет количества
            aggregated_data['Количество'] = aggregated_data['Общее количество'] // total_quantity
            aggregated_data['Общее количество'] = aggregated_data['Количество'] * total_quantity
            aggregated_data['Название сделки'] = deal_name

            # Сохранение данных
            aggregated_data.to_excel(aggregated_file_name, index=False)
            print(f"Итоговые данные сохранены в файл: {aggregated_file_name}")
        except Exception as e:
            print(f"Ошибка при сохранении итоговых данных: {e}")
    


    def save_grouped_data(self, main_file, output_directory, deal_name, total_quantity):
        aggregated_data = self.aggregate_details()

        grouped_data = {}
        for _, row in aggregated_data.iterrows():
            note = row['Примечание']
            filename = self.get_filename_from_note(note, os.path.splitext(os.path.basename(main_file))[0])
            file_path = os.path.join(output_directory, filename)

            if file_path not in grouped_data:
                grouped_data[file_path] = []
            grouped_data[file_path].append(row)

        log_message(f"Grouped data prepared: {grouped_data.keys()}")

        for file_path, data_rows in grouped_data.items():
            df_to_save = pd.DataFrame(data_rows)
            df_to_save['Количество'] = df_to_save['Общее количество'] // total_quantity
            df_to_save = df_to_save[['Обозначение', 'Наименование', 'Примечание', 'Материалы', 'Общее количество', 'Количество']]
            df_to_save.to_excel(file_path, index=False)
            log_message(f"Grouped file saved: {file_path}")





    def save_1c_data(self, aggregated_data, file_path_1c, total_quantity):
        try:
            # Добавляем обычное количество
            aggregated_data['Количество'] = aggregated_data['Общее количество'] // total_quantity

            # Формируем DataFrame для 1С с нужными столбцами
            data_1c = aggregated_data[['Обозначение', 'Наименование', 'Примечание', 'Материалы', 'Общее количество', 'Количество']]
            data_1c.to_excel(file_path_1c, index=False)
            print(f"Итоговый файл 1С успешно сохранен: {file_path_1c}")
        except Exception as e:
            print(f"Ошибка при сохранении итогового файла 1С: {e}")


  
    
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


    def match_material(self, note, materials_dict):
        # Обработка случаев, когда note не является строкой
        if not isinstance(note, str):
            if pd.isna(note):  # Если note равно NaN
                note = ''  # Заменяем на пустую строку
            else:
                note = str(note)  # Преобразуем в строку

        # Поиск кодов материалов и их соответствия
        try:
            nf_codes = re.findall(r'НФ-\d{8}', note)
            materials = [materials_dict.get(code, "Неизвестный материал") for code in nf_codes]
            return ', '.join(filter(None, materials))
        except Exception as e:
            print(f"Ошибка при обработке примечания '{note}': {e}")
            return "Ошибка материала"


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
        self.geometry("400x300")
        self.materials_dict = None
        self.output_directory = None  # Добавлено для сохранения пути выходной директории
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
        #self.text = tk.Text(self, height=30, width=90)
        #self.text.pack(pady=20)

    def choose_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                total_quantity_str = self.quantity_entry.get()
                if not total_quantity_str.isdigit() or not total_quantity_str:
                    raise ValueError("Введенное количество не является допустимым числом или пусто.")
                total_quantity = int(total_quantity_str)

                deal_name = self.deal_entry.get()
                if not deal_name:
                    raise ValueError("Необходимо указать название сделки.")

                if not self.materials_dict:
                    raise ValueError("Словарь материалов не загружен.")

                # Задаем путь для выходной директории
                self.output_directory = os.path.dirname(file_path)

                process_assembly_unit(file_path, total_quantity, self.materials_file_path, deal_name, self.output_directory)

                messagebox.showinfo("Успех", "Файлы обработаны и успешно сохранены!")
            except ValueError as ve:
                messagebox.showerror("Ошибка ввода", str(ve))
            except Exception as e:
                traceback.print_exc()
                messagebox.showerror("Ошибка", str(e))







    # Inside the Application class's choose_materials_file method
    def choose_materials_file(self):
        materials_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if materials_file_path:
            try:
                self.materials_dict = load_materials_data(materials_file_path)
                self.materials_file_path = materials_file_path  # Добавляем эту строку
                messagebox.showinfo("File Chosen", f"Materials file chosen: {materials_file_path}")
                print("Materials Dictionary:", self.materials_dict)
            except Exception as e:
                traceback.print_exc()
                messagebox.showerror("Error", str(e))






if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Обработчик сборочных единиц')
    parser.add_argument('--deal_name', type=str, help='Название сделки')
    parser.add_argument('--quantity', type=int, help='Количество изделий в сделке')
    parser.add_argument('--main_file', type=str, help='Путь к главному файлу')
    parser.add_argument('--materials_file', type=str, help='Путь к файлу материалов 1С')

    args = parser.parse_args()

    # Если аргументы были предоставлены, запускаем логику командной строки
    if args.deal_name and args.quantity and args.main_file and args.materials_file:
        try:
            process_assembly_unit(args.main_file, args.quantity, args.materials_file, args.deal_name)
        except Exception as e:
            log_message(f"Ошибка при обработке файла: {e}")
            traceback.print_exc()
    else:
        # Если аргументы не были предоставлены, запускаем GUI
        app = Application()
        app.mainloop()
