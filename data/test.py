import os
import sys
import pandas as pd

def read_excel_files(directory):
    # Перебираем все файлы в директории
    for file in os.listdir(directory):
        # Проверяем, является ли файл файлом Excel
        if file.endswith(".xlsx") or file.endswith(".xls"):
            # Построение пути к файлу
            file_path = os.path.join(directory, file)
            # Чтение файла Excel
            df = pd.read_excel(file_path)
            # Вывод названия файла и его содержимого
            print(f"Название файла: {file}")
            print("Содержимое:")
            print(df.to_string())

if __name__ == "__main__":
    # Проверяем, был ли передан параметр
    if len(sys.argv) > 1:
        # Если да, используем его как путь к папке
        directory = sys.argv[1]
    else:
        # Если нет, используем текущий рабочий каталог
        directory = os.getcwd()
    
    # Запускаем функцию чтения файлов с выбранной директорией
    read_excel_files(directory)


