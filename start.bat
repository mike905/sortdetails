@echo off
REM Переход в директорию sortdetails на рабочем столе
cd %USERPROFILE%\Desktop\sortdetails

REM Проверка наличия Git и клонирование репозитория, если он еще не склонирован
if not exist "%USERPROFILE%\Desktop\sortdetails\.git" (
  echo Клонирование репозитория...
  git clone https://github.com/mike905/sortdetails %USERPROFILE%\Desktop\sortdetails
)

REM Создание виртуального окружения, если оно еще не создано
if not exist "venv" (
  echo Создание виртуального окружения...
  python -m venv venv
)

REM Активация виртуального окружения
call venv\Scripts\activate

REM Установка зависимостей, если есть файл requirements.txt
if exist "requirements.txt" (
  echo Установка зависимостей...
  pip install -r requirements.txt
)

REM Запуск скрипта (замените script.py на фактическое имя скрипта)
echo Запуск скрипта...
python script.py

pause

