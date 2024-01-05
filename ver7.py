#рассчет + дерево
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
    note_to_filename = {
        # ... Код для примечаний ...
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

def process_sub_files(sub_files_data, main_label, level=0, file=None, parent_qty=1):
    all_details = []
    for _, row in sub_files_data.iterrows():
        filename = row[3] + '.xlsx'
        original_qty = row[5]  # Оригинальное количество без умножения
        qty = row[5] * parent_qty  # умножение количества на родительское количество
        sub_df = read_file(filename)
        if sub_df is not None:
            if level > 0:  # Если это не корневой уровень, то добавляем название сборочной единицы
                file.write("  " * (level - 1) + f"{row[3]} (Сборочная единица на уровне {level}):\n")
            sub_files = extract_section(sub_df, 'Сборочные единицы', 'Детали')
            details = extract_section(sub_df, 'Детали')
            if not details.empty:
                details[5] = qty  # Обновляем количество с учетом родительского
                all_details.append(details)
                print_and_save_tree_data(details, level, main_label, file, original_qty)

            if not sub_files.empty:
                all_details.append(process_sub_files(sub_files, main_label, level + 1, file, qty))

    return pd.concat(all_details) if all_details else pd.DataFrame()

def print_and_save_tree_data(data, level, main_label, file=None, original_qty=None):
    prefix = "  " * level
    for _, row in data.iterrows():
        line = f"{prefix}- {row[3]}: {row[4]}, Оригинальное Количество: {original_qty}, Итоговое Количество: {row[5]}, Примечание: {row[6]}"
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
            combined_details = pd.concat([details_from_main, process_sub_files(sub_files_data, filename.split('.')[0], 0, file)])

            print("Иерархия сборочных единиц и деталей:", file=file)
            print_and_save_tree_data(details_from_main, 0, filename.split('.')[0], file)
            process_sub_files(sub_files_data, filename.split('.')[0], 1, file)

            if not combined_details.empty:
                grouped_details = combined_details.groupby([3, 4, 6])[5].sum().reset_index()
                sorted_details = grouped_details.sort_values(by=[6])

                print(sorted_details)
                sorted_details.to_excel('итоговые_данные.xlsx', index=False)
                print("Итоговые данные сохранены в файл 'итоговые_данные.xlsx'.")

            else:
                print("Нет данных для сортировки и сохранения.")

if __name__ == "__main__":
    main()

