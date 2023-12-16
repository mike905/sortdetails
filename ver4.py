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
        return df
    except FileNotFoundError:
        print(f"Ошибка: Файл '{filename}' не найден.")
        return None

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

def print_and_save_tree_data(data, level, file=None):
    prefix = "  " * level
    for _, row in data.iterrows():
        line = f"{prefix}{row[3]}: {row[4]}, Количество: {row[5]}, Примечание: {row[6]}"
        print(line)
        if file:
            file.write(line + "\n")

def process_sub_files(sub_files_data, level=0, file=None):
    all_details = []
    for _, row in sub_files_data.iterrows():
        filename = row[3] + '.xlsx'
        qty = row[5]
        sub_df = read_file(filename)
        if sub_df is not None:
            sub_files = extract_section(sub_df, 'Сборочные единицы', 'Детали')
            details = extract_section(sub_df, 'Детали')
            if not details.empty:
                details[5] *= qty
                all_details.append(details)
            if not sub_files.empty:
                print_and_save_tree_data(sub_files, level, file)
                all_details.append(process_sub_files(sub_files, level + 1, file))
    return pd.concat(all_details) if all_details else pd.DataFrame()

def main():
    files = list_xlsx_files()
    choice = int(input("Введите номер файла: ")) - 1
    filename = files[choice]

    main_df = read_file(filename)
    if main_df is not None:
        with open("tree_output.txt", "w") as file:
            sub_files_data = extract_section(main_df, 'Сборочные единицы', 'Детали')
            details = extract_section(main_df, 'Детали')
            combined_details = pd.concat([details, process_sub_files(sub_files_data, file=file)])

            if not combined_details.empty:
                # Группировка по обозначению, наименованию и примечанию, суммирование количества
                grouped_details = combined_details.groupby([3, 4, 6])[5].sum().reset_index()

                # Сортировка по примечанию
                sorted_details = grouped_details.sort_values(by=[6])

                print(sorted_details)
                sorted_details.to_excel('итоговые_данные.xlsx', index=False)
                print("Итоговые данные сохранены в файл 'итоговые_данные.xlsx'.")
            else:
                print("Нет данных для сортировки и сохранения.")

if __name__ == "__main__":
    main()

