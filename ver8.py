#gui
import tkinter as tk
from tkinter import filedialog, Text, Label, Frame
import os

# Функции приложения
def choose_file():
    filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Выберите файл",
                                          filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
    if filename:  # Если файл был выбран, начать его обработку
        process_file(filename)

def process_file(filename):
    # Здесь должен быть код для обработки файла
    # Пример обновления лога:
    log_text.configure(state='normal')
    log_text.insert(tk.END, f"Обработка файла: {filename}\n")
    log_text.configure(state='disabled')

# Настройки GUI
root = tk.Tk()
root.title("Программа расчета деталей")

# Верхняя панель с логотипом и слоганом
top_frame = Frame(root)
top_frame.pack(side=tk.TOP, fill=tk.X)
logo_label = Label(top_frame, text="Логотип Компании", bg="blue", fg="white")  # Место для логотипа
logo_label.pack(side=tk.LEFT, padx=10)
slogan_label = Label(top_frame, text="Слоган Компании", bg="green", fg="white")  # Место для слогана
slogan_label.pack(side=tk.LEFT, padx=10)

# Основная область - выбор файла и логи
main_frame = Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True)
choose_file_button = tk.Button(main_frame, text="Выбрать файл", command=choose_file)
choose_file_button.pack(pady=20)
log_text = Text(main_frame, height=15)
log_text.pack(fill=tk.BOTH, expand=True)

# Нижняя панель с информацией о версии
bottom_frame = Frame(root)
bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)
version_label = Label(bottom_frame, text="mkipov 2023 ver 1.0", fg="grey")
version_label.pack(side=tk.RIGHT, padx=10)

root.mainloop()

