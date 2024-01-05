import os
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

# Замените 'your_directory_path' на путь к вашей папке с файлами Excel
read_excel_files('ИНРТ.100.00.00.000')
#read_excel_files('itog/start/ИНРТ.100.00.00.000 Перемешиватель старт')
#read_excel_files('itog/start/ИНРТ.100.00.00.000 Перемешиватель')
#read_excel_files('data')
