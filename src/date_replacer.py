"""
Модуль для поиска и замены дат в текстах.
"""

import re
import logging
from typing import Optional, Tuple, Pattern

logger = logging.getLogger(__name__)


class DateReplacer:
    """Класс для поиска и замены даты."""

    def __init__(self, old_date: str, new_date: str):
        """
        Инициализация заменщика дат.

        Args:
            old_date: Дата, которую ищем (например, "«29» января 2026 г.")
            new_date: Дата, на которую меняем (например, "«26» февраля 2026 г.")
        """
        self.old_date = old_date.strip()
        self.new_date = new_date.strip()

        logger.info(f"Инициализация: Ищем '{self.old_date}', меняем на '{self.new_date}'")

        # Создаем динамический Regex для поиска на основе old_date
        self._DATE_PATTERN: Pattern[str] = self._compile_regex()

    def _compile_regex(self) -> Pattern[str]:
        """Создает гибкое регулярное выражение на основе old_date."""
        # Разбиваем на слова/числа
        tokens = re.findall(r'\S+', self.old_date)
        pattern_parts = [re.escape(t) for t in tokens]

        # Соединяем через \s* (ноль или более пробельных символов)
        pattern_str = r'\s*'.join(pattern_parts)

        # Также добавим опциональную точку в конце, если ищем "г" или "г."
        if pattern_str.endswith('г'):
            pattern_str += r'\.?'

        logger.debug(f"Создан Regex паттерн: {pattern_str}")
        return re.compile(pattern_str)

    def search_pattern(self) -> Pattern[str]:
        """Публичный геттер для regex-паттерна."""
        return self._DATE_PATTERN

    def find_date(self, text: str) -> Optional[str]:
        """Поиск старой даты в тексте."""
        match = self._DATE_PATTERN.search(text)
        if match:
            logger.debug(f"Найдена дата в тексте: '{match.group(0)}'")
            return match.group(0)
        return None

    def replace_date(self, text: str) -> Tuple[str, bool]:
        """Замена даты в тексте на новую."""
        new_text, count = self._DATE_PATTERN.subn(self.new_date, text)
        if count > 0:
            logger.info(f"Заменено вхождений: {count}")
        return (new_text, count > 0)
