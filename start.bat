@echo off
echo Настройка и запуск проекта sortdetails...

:: Проверка наличия Git
where git >nul 2>&1
if %errorlevel% neq 0 (
    echo Git не найден. Пожалуйста, скачайте и установите Git с https://git-scm.com/downloads и запустите этот скрипт снова.
    exit /b
)

:: Проверка наличия Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python не найден. Пожалуйста, скачайте и установите Python с https://www.python.org/downloads/ и запустите этот скрипт снова.
    exit /b
)

:: Клонирование репозитория
echo Клонирование репозитория...
git clone https://github.com/mike905/sortdetails.git
cd sortdetails

:: Создание виртуального окружения
echo Создание виртуального окружения...
python -m venv venv
call venv\Scripts\activate.bat

:: Установка зависимостей
echo Установка зависимостей...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo Проект готов к запуску.
echo Для запуска используйте 'python ваш_скрипт.py' внутри виртуального окружения.
pause

