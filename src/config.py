"""
Конфигурация проекта Docx Date Replacer.
Содержит только неизменяемые константы.
Пути задаются через GUI.
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Config:
    """Неизменяемая конфигурация проекта."""

    # Даты по умолчанию
    old_date: str = "«29» января 2026 г."
    new_date: str = "«26» февраля 2026 г."

    # Настройки обработки
    exclude_prefix: str = "~$"
    file_extension: str = ".docx"
    first_page_paragraphs: int = 50


# Экземпляр по умолчанию
DEFAULT_CONFIG = Config()
