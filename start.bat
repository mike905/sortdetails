@echo off
SET projectFolder=%userprofile%\Desktop\sortdetails

echo Проверка наличия Git и Python...
where git >nul 2>&1
if %errorlevel% neq 0 (
    echo Git не обнаружен. Пожалуйста, скачайте и установите Git с https://git-scm.com/downloads
    exit /b
)

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python не обнаружен. Пожалуйста, скачайте и установите Python с https://www.python.org/downloads/
    exit /b
)

echo Клонирование репозитория на рабочий стол...
git clone https://github.com/mike905/sortdetails.git "%projectFolder%"
cd "%projectFolder%"

echo Создание виртуального окружения...
python -m venv venv
call venv\Scripts\activate.bat

echo Установка зависимостей...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo Проект готов к запуску.
echo Для запуска вашего скрипта, используйте команду: python имя_вашего_скрипта.py
pause

