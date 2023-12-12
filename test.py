import pandas as pd

def print_excel_data(filename):
    try:
        # Чтение всего файла
        data = pd.read_excel(filename)
        print(data)
    except Exception as e:
        print(f"Произошла ошибка при чтении файла {filename}: {e}")

# Использование функции
#print_excel_data("ИЗПА.5177.xlsx")

print_excel_data("итоговые_данные_10.12.23.xlsx")

