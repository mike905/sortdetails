#сумма плюс разделение по файлам примечаниям плюс древовидный вывод

import pandas as pd
import os

def list_xlsx_files():
    files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    for i, file in enumerate(files):
        print(f"{i + 1}: {file}")
    return files

def read_file(filename):
    try:
        df = pd.read_excel(filename, header=None)
        print(f"Файл '{filename}' успешно прочитан.")
        print(df)
        return df
    except FileNotFoundError:
        print(f"Ошибка: Файл '{filename}' не найден.")
        return None

def get_filename_from_note(note, main_filename):
    # Словарь для соответствия примечаний и расшифровок
    note_to_filename = {
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
    base_filename = main_filename.split('.')[0]
    if "НФ-" in note:
        return f"{base_filename} – 1с.xlsx"
    for key in note_to_filename:
        if note.startswith(key):
            return f"{base_filename} – {note_to_filename[key]} {note[len(key):]}-ая очередь.xlsx"
    return f"{base_filename} – Другое.xlsx"

def save_files_by_notes(sorted_details, main_filename):
    for note, group in sorted_details.groupby(6):
        filename = get_filename_from_note(note, main_filename)
        group.to_excel(filename, index=False)
        print(f"Файл '{filename}' сохранен.")

def extract_section(df, start_label, end_label=None):
    try:
        start_idx = df[df.iloc[:, 4] == start_label].index[0] + 1
        if end_label:
            end_idx = df.index[df.iloc[:, 4] == end_label][0]
            section = df.iloc[start_idx:end_idx]
        else:
            section = df.iloc[start_idx:]
        return section.dropna(subset=[5])
    except IndexError:
        print(f"Секция '{start_label}' не найдена.")
        return pd.DataFrame()

def process_sub_files(sub_files_data, qty_multiplier=1, file=None):
    all_details = []
    for _, row in sub_files_data.iterrows():
        filename = row[3] + '.xlsx'
        qty = row[5] * qty_multiplier
        sub_df = read_file(filename)
        if sub_df is not None:
            sub_files = extract_section(sub_df, 'Сборочные единицы', 'Детали')
            details = extract_section(sub_df, 'Детали')
            if not details.empty:
                details[5] *= qty
                all_details.append(details)
            if not sub_files.empty:
                all_details.append(process_sub_files(sub_files, qty, file=file))
    return pd.concat(all_details) if all_details else pd.DataFrame()


def print_tree_data(sub_files_data, details_from_main, file=None):
    for _, row in details_from_main.iterrows():
        line = f"- {row[3]}: {row[4]}, Количество: {row[5]}, Примечание: {row[6]}"
        print(line)
        if file:
            file.write(line + "\n")

    for _, row in sub_files_data.iterrows():
        line = f"- {row[3]}: {row[4]}, Количество: {row[5]}, Примечание: {row[6]}"
        print(line)
        if file:
            file.write(line + "\n")

def main():
    files = list_xlsx_files()
    choice = int(input("Введите номер файла: ")) - 1
    filename = files[choice]

    main_df = read_file(filename)
    if main_df is not None:
        with open("tree_output.txt", "w") as file:
            sub_files_data = extract_section(main_df, 'Сборочные единицы', 'Детали')
            details_from_main = extract_section(main_df, 'Детали')
            combined_details = pd.concat([details_from_main, process_sub_files(sub_files_data, 1, file=file)])


            print_tree_data(sub_files_data, details_from_main, file=file)

            if not combined_details.empty:
                grouped_details = combined_details.groupby([3, 4, 6])[5].sum().reset_index()
                sorted_details = grouped_details.sort_values(by=[6])

                print(sorted_details)
                sorted_details.to_excel('итоговые_данные.xlsx', index=False)
                print("Итоговые данные сохранены в файл 'итоговые_данные.xlsx'.")
                save_files_by_notes(sorted_details, filename)
            else:
                print("Нет данных для сортировки и сохранения.")

if __name__ == "__main__":
    main()




