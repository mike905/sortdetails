import pandas as pd

def read_main_file(filename):
    try:
        # Чтение файла
        df = pd.read_excel(filename, header=None)
        print("Файл успешно прочитан.")

        # Находим индекс строки, содержащей "Сборочные единицы"
        start_idx = df[df.iloc[:, 4] == 'Сборочные единицы'].index[0] + 1
        print("Найдена строка 'Сборочные единицы'.")

        # Находим индекс строки, содержащей "Детали"
        end_idx = df[df.iloc[:, 4] == 'Детали'].index[0]
        print("Найдена строка 'Детали'.")

        # Извлечение данных о вложенных файлах, игнорируем пустые строки
        sub_files_data = df.iloc[start_idx:end_idx].dropna(subset=[3, 4, 5])
        print("Данные о вложенных файлах извлечены.")

        return sub_files_data
    except Exception as e:
        print(f"Произошла ошибка при обработке файла {filename}: {e}")
        return None

def read_sub_files(sub_files_data):
    all_details = []

    for _, row in sub_files_data.iterrows():
        sub_filename = row[3] + '.xlsx'
        qty = row[5]

        try:
            # Чтение вложенного файла
            sub_df = pd.read_excel(sub_filename, header=None)
            print(f"Файл {sub_filename} успешно прочитан.")

            # Проверяем наличие строки 'Детали'
            if 'Детали' in sub_df.iloc[:, 4].values:
                start_idx = sub_df[sub_df.iloc[:, 4] == 'Детали'].index[0] + 1

                # Извлечение данных, начиная после строки 'Детали'
                details = sub_df.iloc[start_idx:, :7].dropna(subset=[3, 4, 5])
                details[5] *= qty  # Умножение количества на значение из основного файла
                all_details.append(details)
            else:
                print(f"В файле {sub_filename} не найдена строка 'Детали'.")
        except Exception as e:
            print(f"Произошла ошибка при обработке файла {sub_filename}: {e}")

    # Объединение всех данных деталей
    if all_details:
        combined_details = pd.concat(all_details)
        return combined_details
    else:
        print("Нет данных для объединения.")
        return None



def main():
    main_filename = input("Введите название основного файла (с расширением): ")
    sub_files_data = read_main_file(main_filename)

    if sub_files_data is not None:
        combined_details = read_sub_files(sub_files_data)

        if combined_details is not None:
            # Сортировка по примечанию
            sorted_details = combined_details.sort_values(by=6)
            print("Данные успешно отсортированы.")

            # Вывод в консоль и сохранение в файл
            print(sorted_details)
            sorted_details.to_excel('итоговые_данные.xlsx', index=False)
            print("Итоговые данные сохранены в файл 'итоговые_данные.xlsx'.")
        else:
            print("Нет данных для сортировки и сохранения.")
    else:
        print("Нет данных для обработки из основного файла.")

if __name__ == "__main__":
    main()


