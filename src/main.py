"""
Точка входа приложения Docx Date Replacer.

Замена даты в блоке «УТВЕРЖДАЮ» в документах .docx.
"""

import os
import sys
from datetime import datetime

# Добавляем корневую папку в путь
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.config import (
    SOURCE_FOLDER,
    OUTPUT_FOLDER,
    NEW_DATE,
    EXCLUDE_PREFIX,
    FILE_EXTENSION,
)
from src.date_replacer import DateReplacer
from src.docx_processor import DocxProcessor


def print_header():
    """Вывод заголовка программы."""
    print("=" * 60)
    print("Docx Date Replacer - Замена даты в документах .docx")
    print("=" * 60)
    print()


def print_config():
    """Вывод конфигурации."""
    print("Конфигурация:")
    print(f"  Исходная папка: {SOURCE_FOLDER}")
    print(f"  Папка вывода:   {OUTPUT_FOLDER}")
    print(f"  Новая дата:     {NEW_DATE}")
    print()


def print_statistics(
    total_files: int,
    processed_with_changes: int,
    processed_without_changes: int,
    total_replacements: int,
    errors: list,
    elapsed_time: float,
):
    """
    Вывод статистики обработки.

    Args:
        total_files: Всего файлов найдено
        processed_with_changes: Файлов с заменами
        processed_without_changes: Файлов без изменений
        total_replacements: Всего замен выполнено
        errors: Список ошибок
        elapsed_time: Время выполнения в секундах
    """
    print()
    print("=" * 60)
    print("Статистика обработки:")
    print("=" * 60)
    print(f"  Всего файлов найдено:       {total_files}")
    print(f"  Файлов с заменами:          {processed_with_changes}")
    print(f"  Файлов без изменений:       {processed_without_changes}")
    print(f"  Всего замен выполнено:      {total_replacements}")
    print(f"  Время выполнения:           {elapsed_time:.2f} сек.")

    if errors:
        print()
        print("Ошибки:")
        for filepath, error in errors:
            print(f"  ❌ {os.path.basename(filepath)}: {error}")

    print()
    print(f"Результат сохранён в: {OUTPUT_FOLDER}")
    print("=" * 60)


def main():
    """Основная функция."""
    print_header()
    print_config()
    
    # Проверка существования исходной папки
    if not os.path.exists(SOURCE_FOLDER):
        print(f"❌ Ошибка: Исходная папка не найдена: {SOURCE_FOLDER}")
        sys.exit(1)
    
    # Инициализация компонентов
    date_replacer = DateReplacer(NEW_DATE)
    processor = DocxProcessor(date_replacer)
    
    # Поиск файлов
    print("Поиск файлов .docx...")
    docx_files = processor.find_docx_files(SOURCE_FOLDER, EXCLUDE_PREFIX)
    print(f"  Найдено файлов: {len(docx_files)}")
    print()
    
    if not docx_files:
        print("❌ Файлы .docx не найдены")
        sys.exit(1)
    
    # Копирование структуры папок
    print("Создание структуры папок вывода...")
    processor.copy_folder_structure(SOURCE_FOLDER, OUTPUT_FOLDER)
    print(f"  Папка вывода: {OUTPUT_FOLDER}")
    print()
    
    # Обработка файлов
    print("Обработка файлов...")
    print("-" * 60)

    total_files = len(docx_files)
    processed_with_changes = 0
    processed_without_changes = 0
    total_replacements = 0
    errors = []

    start_time = datetime.now()

    for i, filepath in enumerate(docx_files, 1):
        filename = os.path.basename(filepath)
        output_path = processor.get_output_path(filepath, SOURCE_FOLDER, OUTPUT_FOLDER)

        # Обработка документа
        success, message, replacements = processor.process_document(filepath, output_path)

        if success:
            if replacements > 0:
                processed_with_changes += 1
                total_replacements += replacements
                print(f"  [{i}/{total_files}] ✓ {filename}: {message}")
            else:
                processed_without_changes += 1
                print(f"  [{i}/{total_files}] ○ {filename}: {message}")
        else:
            errors.append((filepath, message))
            print(f"  [{i}/{total_files}] ❌ {filename}: {message}")

    end_time = datetime.now()
    elapsed_time = (end_time - start_time).total_seconds()

    # Вывод статистики
    print_statistics(
        total_files,
        processed_with_changes,
        processed_without_changes,
        total_replacements,
        errors,
        elapsed_time,
    )
    
    # Итоговый статус
    if processed_with_changes > 0:
        print("✅ Обработка завершена успешно!")
    else:
        print("⚠️  Обработка завершена без изменений")
        print("   Возможно, в файлах нет даты «29» января 2026 г.")


if __name__ == "__main__":
    main()
