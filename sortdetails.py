import os
import pandas as pd

def list_excel_files():
    # Получение списка всех файлов Excel в текущей директории
    return [f for f in os.listdir('.') if f.endswith('.xlsx')]

def read_file(filename):
    # Функция для чтения файла
    try:
        df = pd.read_excel(filename, header=None)
        print(f"Содержимое файла {filename}:")
        print(df)
        return df
    except Exception as e:
        print(f"Произошла ошибка при чтении файла {filename}: {e}")
        return pd.DataFrame()


def process_files(main_details, processed_files=set(), level=0, parent_quantities=None):
    all_details = []
    quantities = {}

    # Обработка текущего файла
    try:
        # Проверка наличия строки "Сборочные единицы"
        if 'Сборочные единицы' in main_details.iloc[:, 4].values:
            start_idx_assembly = main_details[main_details.iloc[:, 4] == 'Сборочные единицы'].index[0] + 1
        else:
            start_idx_assembly = None

        # Проверка наличия строки "Детали"
        if 'Детали' in main_details.iloc[:, 4].values:
            end_idx_assembly = main_details[main_details.iloc[:, 4] == 'Детали'].index[0]
        else:
            # Если строка "Детали" отсутствует, устанавливаем end_idx_assembly в конец документа
            end_idx_assembly = len(main_details)
    except Exception as e:
        print("  " * level + f"Ошибка при обработке индексов: {e}")
        return pd.DataFrame()

    # Извлечение данных о сборочных единицах
    if start_idx_assembly is not None and end_idx_assembly is not None:
        sub_files_data = main_details.iloc[start_idx_assembly:end_idx_assembly]
        sub_files = sub_files_data.iloc[:, 3].dropna().unique()
        sub_files = [f'{name}.xlsx' for name in sub_files if f'{name}.xlsx' in os.listdir('.')]

        # Сохранение количества каждой сборочной единицы
        for _, row in sub_files_data.iterrows():
            if pd.notna(row[5]):  # Проверка на наличие NaN
                quantities[row[3]] = int(row[5])
            else:
                quantities[row[3]] = 1  # Установка значения по умолчанию для NaN

    else:
        sub_files = []

    # Обработка вложенных файлов
    for sub_filename in sub_files:
        if sub_filename in processed_files:
            continue

        print("  " * level + f"Обработка файла: {sub_filename}")
        details = read_file(sub_filename)
        processed_files.add(sub_filename)

        if not details.empty:
            all_details.extend(process_files(details, processed_files, level + 1, quantities))

    # Добавление деталей из текущего файла
    if end_idx_assembly is not None:
        detail_rows = main_details.iloc[end_idx_assembly + 1:].dropna(subset=[3, 4, 5])
        if isinstance(detail_rows, pd.DataFrame):
            # Умножение количества деталей на количество соответствующих сборочных единиц
            for idx, row in detail_rows.iterrows():
                item_quantity = int(row[5])  # Преобразование количества в int
                if row[3] in quantities:
                    item_quantity *= quantities[row[3]]

                detail_rows.at[idx, 5] = item_quantity

            all_details.append(detail_rows)

            # Вывод информации о деталях в сборочной единице
            print("  " * level + f"Детали в сборочной единице {main_details.iloc[0, 3]}:")
            for _, row in detail_rows.iterrows():
                print("  " * level + f"Обозначение: {row[3]}, Наименование: {row[4]}, Количество: {row[5]}")

    return pd.concat(all_details, ignore_index=True) if all_details else pd.DataFrame()





def main():
    files = list_excel_files()
    if not files:
        print("В текущей директории нет файлов Excel.")
        return

    print("Выберите главный файл, указав соответствующий номер:")
    for i, file in enumerate(files, 1):
        print(f"{i}. {file}")

    choice = int(input("Введите номер файла: ")) - 1
    if choice < 0 or choice >= len(files):
        print("Неверный выбор.")
        return

    main_filename = files[choice]
    main_details = read_file(main_filename)
    if main_details.empty:
        print("Главный файл пустой.")
        return

    combined_details = process_files(main_details)

    if combined_details.empty:
        print("Нет данных для обработки.")
        return

    # Вывод итоговых данных в консоль и сохранение в файл
    print("Итоговые данные:")
    print(combined_details)

    # Сохранение итоговых данных в файл
    combined_details.columns = main_details.iloc[0, :].values
    combined_details.to_excel('итоговые_данные.xlsx', index=False)
    print("Итоговые данные сохранены в файл 'итоговые_данные.xlsx'.")

if __name__ == "__main__":
    main()

