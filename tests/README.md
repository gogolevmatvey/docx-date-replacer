# Тестирование проекта Docx Date Replacer

## Запуск тестов

### Все тесты
```powershell
py -m unittest discover -s tests
```

### Конкретный модуль
```powershell
# Тесты date_replacer
py -m unittest tests.test_date_replacer

# Тесты docx_processor
py -m unittest tests.test_docx_processor

# Интеграционные тесты
py -m unittest tests.test_integration
```

### Конкретный тест
```powershell
py -m unittest tests.test_date_replacer.TestDateReplacer.test_find_date_standard_format
```

### С подробным выводом
```powershell
py -m unittest discover -s tests -v
```

## Структура тестов

```
tests/
├── __init__.py
├── test_date_replacer.py      # Юнит-тесты для DateReplacer
├── test_docx_processor.py     # Юнит-тесты для DocxProcessor
└── test_integration.py        # Интеграционные тесты
```

## Покрытие тестами

### DateReplacer
- ✅ Поиск даты в различных форматах
- ✅ Замена даты
- ✅ Валидация формата даты
- ✅ Проверка блока «УТВЕРЖДАЮ»

### DocxProcessor
- ✅ Загрузка/сохранение документов
- ✅ Обработка параграфов
- ✅ Обработка таблиц
- ✅ Поиск файлов .docx
- ✅ Копирование структуры папок

### Интеграционные тесты
- ✅ Полный цикл обработки
- ✅ Документы с таблицами
- ✅ Документы без даты
- ✅ Несколько дат в документе
- ✅ Сохранение структуры папок

## Требования

Для запуска тестов требуется:
```
python-docx>=1.2.0
```

Установка:
```powershell
pip install -r requirements.txt
```
