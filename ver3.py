import pandas as pd
import os

# Функция для листинга xlsx файлов в текущей директории
def list_xlsx_files():
    files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    for i, file in enumerate(files):
        print(f"{i + 1}: {file}")
    return files

# Функция для чтения Excel файла
def read_file(filename):
    try:
        df = pd.read_excel(filename, header=None)
        return df
    except FileNotFoundError:
        print(f"Ошибка: Файл '{filename}' не найден.")
        return None

# Функция для извлечения секции из файла
def extract_section(df, start_label, end_label=None):
    try:
        start_idx = df[df.iloc[:, 4] == start_label].index[0] + 1
        if end_label is None:
            section = df.iloc[start_idx:]
        else:
            end_idx = df[df.iloc[:, 4] == end_label].index[0]
            section = df.iloc[start_idx:end_idx]
        section = section.dropna(subset=[5])
        return section
    except IndexError:
        return pd.DataFrame()

# Функция для обработки вложенных файлов
def process_sub_files(sub_files_data, level=0):
    all_details = []
    for _, row in sub_files_data.iterrows():
        sub_filename = str(row[3]) + '.xlsx'
        qty = row[5]
        sub_df = read_file(sub_filename)
        if sub_df is not None:
            # Обработка деталей в файле
            details = extract_section(sub_df, 'Детали')
            if not details.empty:
                details[5] *= qty
                all_details.append(details)
                print(f"{'  ' * level}- {row[3]}: {row[4]}, Количество: {qty}, Примечание: {row[6]}")
                for _, detail_row in details.iterrows():
                    print(f"{'  ' * (level + 1)}- {detail_row[3]}: {detail_row[4]}, Количество: {detail_row[5]}, Примечание: {detail_row[6]}")
    return pd.concat(all_details) if all_details else None

# Основная функция
def main():
    files = list_xlsx_files()
    choice = int(input("Введите номер файла: ")) - 1
    main_filename = files[choice]

    main_df = read_file(main_filename)
    if main_df is not None:
        sub_files_data = extract_section(main_df, 'Сборочные единицы', 'Детали')
        details_from_main = extract_section(main_df, 'Детали')
        combined_details = pd.concat([details_from_main, process_sub_files(sub_files_data)])
        if not combined_details.empty:
            sorted_details = combined_details.sort_values(by=[3, 6])
            print("Итоговые данные:")
            print(sorted_details)
            sorted_details.to_excel('итоговые_данные.xlsx', index=False)
            print("Файл 'итоговые_данные.xlsx' сохранен.")
        else:
            print("Нет данных для сортировки и сохранения.")
    else:
        print("Основной файл не содержит данных.")

if __name__ == "__main__":
    main()

