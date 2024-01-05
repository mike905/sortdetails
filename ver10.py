#gui chat gpt

import tkinter as tk
from tkinter import filedialog, Listbox, Text, Button, Frame, Label
from PIL import Image, ImageTk
import datetime
import json
import os

# Файл для сохранения истории сессий
session_log_file = "session_history.json"

# Путь к логотипу (убедитесь, что файл находится в той же директории, что и скрипт)
logo_path = "inerta.jpeg"

def choose_file():
    filename = filedialog.askopenfilename(initialdir="/", title="Выберите файл",
                                          filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
    if filename:
        add_session_log(filename)
        process_file(filename)

def process_file(filename):
    # Здесь должен быть код для обработки файла
    log_text.configure(state='normal')
    log_text.insert(tk.END, f"Обработка файла: {filename}\n")
    log_text.configure(state='disabled')

def add_session_log(filename):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    session_info = f"{now}: {filename}"
    sessions_list.insert(tk.END, session_info)
    save_sessions_to_file(session_info)

def save_sessions_to_file(session_info):
    # Загрузить существующие сессии, если они есть
    if os.path.exists(session_log_file):
        with open(session_log_file, 'r') as file:
            sessions = json.load(file)
    else:
        sessions = []

    sessions.append(session_info)
    with open(session_log_file, 'w') as file:
        json.dump(sessions, file)

def load_sessions():
    if os.path.exists(session_log_file):
        with open(session_log_file, 'r') as file:
            sessions = json.load(file)
            for session in sessions:
                sessions_list.insert(tk.END, session)

# Настройки GUI
root = tk.Tk()
root.title("Chat-Style File Processor")

# Панель сессий слева
session_frame = tk.Frame(root, bd=2, relief="sunken")
session_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
sessions_label = tk.Label(session_frame, text="Сессии", font=("Helvetica", 16))
sessions_label.pack(side=tk.TOP, fill=tk.X)
sessions_list = Listbox(session_frame)
sessions_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
load_sessions()  # Загрузить историю сессий при запуске

# Верхняя панель для логотипа
top_frame = tk.Frame(root, bd=2, relief="sunken")
top_frame.pack(side=tk.TOP, fill=tk.X)
logo_image = Image.open(logo_path)  # Откройте изображение с помощью Pillow
logo_image = logo_image.resize((260, 50))  # Масштабируйте или измените размер, если необходимо
logo_img = ImageTk.PhotoImage(logo_image)  # Конвертируйте в PhotoImage
logo_label = tk.Label(top_frame, image=logo_img)
logo_label.pack(side=tk.LEFT, padx=10)

# Основное поле чата для логов
chat_frame = tk.Frame(root, bd=2, relief="sunken")
chat_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
log_text = tk.Text(chat_frame, height=20)
log_text.pack(fill=tk.BOTH, expand=True)

# Кнопка выбора файла в нижней части основного поля чата
choose_file_button = tk.Button(chat_frame, text="Выбрать файл", command=choose_file)
choose_file_button.pack(side=tk.BOTTOM)

root.mainloop()

