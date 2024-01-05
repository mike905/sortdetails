#gui+group+itog+tree

import pandas as pd 
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback


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
        # Collecting data for export
        lines = []
        self.collect_tree_data(self.assembly_unit, lines)
        
        # Preparing data
        data = []
        for level, name, quantity, note in lines:
            row = [None] * 11  # Assuming 11 columns as in your example
            row[level + 1] = name
            row[-2] = quantity
            row[-1] = note
            data.append(row)

        # Creating DataFrame with the desired structure
        columns = ['Unnamed: 0', 'Наименование', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Структура изделия', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10']
        df = pd.DataFrame(data, columns=columns)  # Ensure df is defined here

        # Saving DataFrame to Excel
        df.to_excel(filename, index=False)
        print(f"Дерево сборочных единиц сохранено в {filename}")

    def collect_tree_data(self, unit, lines, level=0):
        lines.append((level, unit.filename, unit.quantity, 'Сборочная единица'))
        for detail in unit.details:
            lines.append((level + 1, detail.designation, detail.total_qty, detail.note))
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

    def save_grouped_data(self, main_filename, directory):
        aggregated_data = self.aggregate_details()
        print(f"Начинается сохранение группированных данных в {directory}")
        
        for note, group in aggregated_data.groupby('Примечание'):
            print(f"Обрабатывается примечание: {note}")
            filename = self.get_filename_from_note(note, main_filename)
            file_path = os.path.join(directory, filename)  # Формируем полный путь к файлу
            print(f"Сохранение файла: {file_path}")
            group.to_excel(file_path, index=False)
            print(f"Сохранено {len(group)} строк в файл: {file_path}")

    def get_filename_from_note(self, note, main_filename):
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
            # Дополнительные примечания
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

        base_filename = main_filename.split('.')[0]
        filename_notes = []

        # Разбиваем примечание на отдельные части
        note_parts = note.split(',')
        for part in note_parts:
            if part in note_to_filename:
                filename_notes.append(note_to_filename[part])
            elif "НФ-" in part:  # Обрабатываем специальные случаи для материалов
                # Дополните логику для обработки специальных случаев материалов
                filename_notes.append("Специальный материал")  # Пример, нужно уточнить

        if filename_notes:
            # Соединяем все получившиеся части примечаний для формирования имени файла
            note_str = ' & '.join(filename_notes)
            return f"{base_filename} – {note_str}.xlsx"
        else:
            return f"{base_filename} – Другое.xlsx"

     



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
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                filename_without_extension = os.path.splitext(os.path.basename(file_path))[0]
                directory = filename_without_extension  # папка называется как основной файл

                if not os.path.exists(directory):
                    os.makedirs(directory)  # Создаем папку, если не существует

                # Обработка файла
                main_assembly = AssemblyUnit(filename_without_extension)
                main_assembly.process_file()

                # Вывод дерева и агрегированных данных в текстовый виджет
                data_output = DataOutput(main_assembly)
                self.text.insert(tk.END, "Иерархия сборочных единиц и деталей:\n")
                data_output.print_tree(target=self.text)

                # Агрегирование и вывод итоговых данных
                aggregator = DataAggregator(main_assembly)
                aggregated_data = aggregator.aggregate_details()  # Получение агрегированных данных
                self.text.insert(tk.END, "\nИтоговые данные:\n")
                aggregator.print_aggregated_data(text_widget=self.text)

                # Сохранение итоговых данных
                aggregated_file_name = os.path.join(directory, f"{filename_without_extension}_aggregated.xlsx")
                aggregator.save_aggregated_data(aggregated_data, aggregated_file_name)

                # Сохранение группированных данных по примечаниям
                aggregator.save_grouped_data(filename_without_extension, directory)

                # Сохранение структуры дерева
                tree_file_name = os.path.join(directory, f"{filename_without_extension}_tree_structure.xlsx")
                data_output.save_tree_to_excel(tree_file_name)

                # Всплывающее сообщение об успешной обработке
                messagebox.showinfo("Обработка завершена", "Данные успешно обработаны и сохранены!")

            except Exception as e:
                traceback.print_exc()  # Печатает полную трассировку ошибки
                messagebox.showerror("Ошибка", str(e))

def main():
    app = Application()
    app.mainloop()

if __name__ == "__main__":
    main()

