@echo off
REM Запуск Docx Date Replacer

REM Активируем виртуальное окружение
call .venv\Scripts\activate.bat

REM Запускаем приложение
python -m src.main

pause
